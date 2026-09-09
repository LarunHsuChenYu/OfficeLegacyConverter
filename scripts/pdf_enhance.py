#!/usr/bin/env python3
"""Extract PDF tables (pdfplumber) and page PNGs (PyMuPDF) for Markdown post-processing.

Also emits per-page quality signals (likely image tables) and optional OCR captions.
"""

from __future__ import annotations

import json
import re
import shutil
import subprocess
import sys
from collections import Counter
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
elif hasattr(sys.stdout, "buffer"):
    import io

    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

CONTENT_IMAGE_MIN_AREA_RATIO = 0.03
FULL_PAGE_IMAGE_AREA_RATIO = 0.90
OCR_TABLE_DPI = 400


def _table_to_markdown(rows: list[list[str]]) -> str:
    if not rows:
        return ""

    normalized: list[list[str]] = []
    for row in rows:
        cells = [str(c or "").replace("\n", " ").strip() for c in row]
        normalized.append(cells)

    if not normalized:
        return ""

    col_count = max(len(r) for r in normalized)
    for row in normalized:
        while len(row) < col_count:
            row.append("")

    def escape(cell: str) -> str:
        return cell.replace("|", "\\|")

    header = normalized[0]
    lines = [
        "| " + " | ".join(escape(c) for c in header) + " |",
        "| " + " | ".join("---" for _ in header) + " |",
    ]
    for row in normalized[1:]:
        lines.append("| " + " | ".join(escape(c) for c in row) + " |")
    return "\n".join(lines)


def _empty_cell_ratio(rows: list[list[str]]) -> float:
    total = 0
    empty = 0
    for row in rows:
        for cell in row:
            total += 1
            if not str(cell or "").strip():
                empty += 1
    return empty / total if total else 0.0


def _is_junk_table(rows: list[list[str]]) -> bool:
    """Detect pdfplumber false positives (e.g. whole slide text in one cell)."""
    if not rows:
        return True
    row_count = len(rows)
    col_count = max((len(r) for r in rows), default=0)
    if row_count < 2:
        return True
    if col_count <= 2 and row_count <= 2:
        # Single fat cell wrapping title + bullets
        nonempty = [str(c or "").strip() for r in rows for c in r if str(c or "").strip()]
        if len(nonempty) <= 2:
            return True
        if nonempty and max(len(c) for c in nonempty) > 120:
            return True
    return False


def _page_text_hints(page) -> tuple[str, list[str]]:
    """Return (title_hint, bullet_hints) from pdfplumber page text."""
    try:
        raw = page.extract_text() or ""
    except Exception:
        raw = ""
    lines = [ln.strip() for ln in raw.splitlines() if ln and ln.strip()]
    title = lines[0] if lines else ""
    bullets: list[str] = []
    for ln in lines[1:]:
        if ln.startswith("➢") or ln.startswith("•") or ln.startswith("-") or ln.startswith("·"):
            bullets.append(re.sub(r"^[➢•·\-]\s*", "", ln).strip())
        elif "共計" in ln or "需換算" in ln or "KPI" in ln:
            bullets.append(ln)
        if len(bullets) >= 4:
            break
    return title, bullets


def _build_image_caption(page_num: int, title: str, bullets: list[str]) -> str:
    parts = [
        f"第 {page_num} 頁影像表格截圖（Excel 篩選列／明細數據）。",
    ]
    if title:
        short = title.replace("\n", " ").strip()
        if len(short) > 80:
            short = short[:80] + "…"
        parts.append(f"標題「{short}」。")
    if bullets:
        joined = "；".join(b[:60].rstrip("。；; ") for b in bullets[:3])
        parts.append(f"頁面文字結論：{joined}。")
    parts.append("自動轉換未能還原儲存格，請對照下方截圖。")
    return "".join(parts)


def _content_image_boxes(page) -> list[tuple[float, float, float, float]]:
    """Return PyMuPDF-coordinate boxes for content images, excluding backgrounds/logos."""
    page_area = float(page.width * page.height)
    if page_area <= 0:
        return []

    boxes = []
    for image in page.images or []:
        width = float(image.get("width") or 0)
        height = float(image.get("height") or 0)
        area_ratio = (width * height) / page_area
        if CONTENT_IMAGE_MIN_AREA_RATIO <= area_ratio < FULL_PAGE_IMAGE_AREA_RATIO:
            boxes.append(
                (
                    float(image.get("x0") or 0),
                    float(page.height - (image.get("y1") or page.height)),
                    float(image.get("x1") or page.width),
                    float(page.height - (image.get("y0") or 0)),
                )
            )
    return boxes


def _find_tesseract() -> str | None:
    configured = shutil.which("tesseract")
    if configured:
        return configured

    script_dir = Path(__file__).resolve().parent
    candidates = [
        script_dir / "tesseract" / "tesseract.exe",
        script_dir.parent / "tools" / "tesseract" / "tesseract.exe",
        Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
        Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
    ]
    return next((str(path) for path in candidates if path.is_file()), None)


def _find_tessdata_dir() -> str | None:
    script_dir = Path(__file__).resolve().parent
    candidates = [
        script_dir / "tessdata",
        script_dir.parent / "tools" / "tesseract" / "tessdata",
    ]
    return next(
        (str(path) for path in candidates if (path / "eng.traineddata").is_file()),
        None,
    )


def _available_ocr_languages(executable: str, tessdata_dir: str | None) -> set[str]:
    command = [executable]
    if tessdata_dir:
        command.extend(["--tessdata-dir", tessdata_dir])
    command.append("--list-langs")
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=15,
        check=False,
    )
    if completed.returncode != 0:
        return set()
    return {
        line.strip()
        for line in completed.stdout.splitlines()
        if line.strip() and not line.lower().startswith("list of available")
    }


def _run_tesseract(
    executable: str,
    tessdata_dir: str | None,
    image_path: Path,
    language: str,
    psm: int,
) -> tuple[str, str]:
    command = [executable]
    if tessdata_dir:
        command.extend(["--tessdata-dir", tessdata_dir])
    command.extend(
        [str(image_path), "stdout", "-l", language, "--psm", str(psm)]
    )
    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=120,
        check=False,
    )
    if completed.returncode != 0:
        detail = completed.stderr.strip().replace("\n", " ")[:240]
        return "", f"error:exit_{completed.returncode}:{detail}"
    cleaned = "\n".join(
        line.strip()
        for line in (completed.stdout or "").splitlines()
        if line.strip()
    )
    return cleaned, "ok" if cleaned else "ok_empty"


def _numeric_tokens(text: str) -> list[str]:
    return [
        token.replace(" ", "")
        for token in re.findall(
            r"(?<![\w])(?:n/?a|[-+]?\d[\d,.]*(?:%|/\d[\d,.]*)?)",
            text,
            flags=re.IGNORECASE,
        )
    ]


def _merge_sparse_numeric_supplement(primary: str, sparse: str) -> str:
    """Append PSM-11 lines only when they add numeric occurrences missed by PSM-3."""
    if not sparse.strip():
        return primary

    primary_counts = Counter(_numeric_tokens(primary))
    observed_counts: Counter[str] = Counter()
    supplements: list[str] = []
    for line in sparse.splitlines():
        tokens = _numeric_tokens(line)
        adds_missing_value = False
        for token in tokens:
            observed_counts[token] += 1
            if observed_counts[token] > primary_counts[token]:
                adds_missing_value = True
        if adds_missing_value:
            supplements.append(line)

    if not supplements:
        return primary
    return (
        f"{primary}\n"
        "--- OCR 補充辨識（稀疏數值） ---\n"
        + "\n".join(supplements)
    )


def _try_ocr_text(image_path: Path) -> tuple[str, str]:
    """Return complete OCR text using table and sparse-text passes. Never raises."""
    executable = _find_tesseract()
    if not executable:
        return "", "skipped_no_executable"

    try:
        tessdata_dir = _find_tessdata_dir()
        languages = _available_ocr_languages(executable, tessdata_dir)
        selected = "+".join(lang for lang in ("chi_tra", "eng") if lang in languages)
        if not selected:
            return "", "skipped_no_language"

        primary, primary_status = _run_tesseract(
            executable, tessdata_dir, image_path, selected, psm=3
        )
        if primary_status.startswith("error", 0):
            return "", primary_status
        sparse, sparse_status = _run_tesseract(
            executable, tessdata_dir, image_path, selected, psm=11
        )
        if sparse_status.startswith("error", 0):
            return primary, "ok" if primary else primary_status
        text = _merge_sparse_numeric_supplement(primary, sparse)
    except Exception as exc:
        return "", f"error:{type(exc).__name__}:{exc}"

    if not text.strip():
        return "", "ok_empty"
    return text, "ok"


def main() -> int:
    if len(sys.argv) != 4:
        print(
            json.dumps(
                {
                    "error": "用法: pdf_enhance.py <input_pdf> <output_dir> <basename>",
                    "pages": [],
                    "tables": [],
                    "image_paths": [],
                },
                ensure_ascii=False,
            )
        )
        return 1

    input_pdf = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])
    basename = sys.argv[3]

    result: dict = {
        "input_pdf": str(input_pdf),
        "basename": basename,
        "images_folder": f"{basename}/",
        "pages": [],
        "tables": [],
        "image_paths": [],
        "errors": [],
        "pages_exported": 0,
        "tables_extracted": 0,
        "tables_extracted_total": 0,
        "tables_extracted_good": 0,
    }

    if not input_pdf.is_file():
        result["error"] = f"找不到 PDF：{input_pdf}"
        print(json.dumps(result, ensure_ascii=False))
        return 1

    images_dir = output_dir / basename
    images_dir.mkdir(parents=True, exist_ok=True)

    page_by_num: dict[int, dict] = {}
    content_boxes_by_page: dict[int, list[tuple[float, float, float, float]]] = {}

    # --- Page PNG export via PyMuPDF ---
    try:
        import fitz  # PyMuPDF

        doc = fitz.open(str(input_pdf))
        for page_idx in range(len(doc)):
            page_num = page_idx + 1
            image_name = f"img{page_num}.png"
            image_path = images_dir / image_name
            page = doc.load_page(page_idx)
            pix = page.get_pixmap(dpi=150)
            pix.save(str(image_path))
            rel_path = f"{basename}/{image_name}"
            marker = f"【img{page_num}】"
            page_info = {
                "index": page_num,
                "filename": image_name,
                "page": page_num,
                "page_number": page_num,
                "marker": marker,
                "image_path": rel_path,
                "char_count": 0,
                "embedded_image_count": 0,
                "content_image_count": 0,
                "good_table_count": 0,
                "junk_table_count": 0,
                "likely_image_table": False,
                "image_caption": "",
                "ocr_preview": "",
                "ocr_status": "not_attempted",
                "context_hints": [],
            }
            result["pages"].append(page_info)
            result["image_paths"].append(rel_path)
            page_by_num[page_num] = page_info
        result["pages_exported"] = len(result["pages"])
        doc.close()
    except ImportError:
        result["errors"].append("PyMuPDF 未安裝（pip install pymupdf），略過頁面 PNG 匯出")
    except Exception as exc:
        result["errors"].append(f"PyMuPDF 頁面匯出失敗：{exc}")

    # --- Table extraction + page signals via pdfplumber ---
    try:
        import pdfplumber

        with pdfplumber.open(str(input_pdf)) as pdf:
            for page_idx, page in enumerate(pdf.pages):
                page_num = page_idx + 1
                page_info = page_by_num.get(page_num)
                if page_info is None:
                    page_info = {
                        "index": page_num,
                        "filename": f"img{page_num}.png",
                        "page": page_num,
                        "page_number": page_num,
                        "marker": f"【img{page_num}】",
                        "image_path": f"{basename}/img{page_num}.png",
                        "char_count": 0,
                        "embedded_image_count": 0,
                        "content_image_count": 0,
                        "good_table_count": 0,
                        "junk_table_count": 0,
                        "likely_image_table": False,
                        "image_caption": "",
                        "ocr_preview": "",
                        "ocr_status": "not_attempted",
                        "context_hints": [],
                    }
                    result["pages"].append(page_info)
                    page_by_num[page_num] = page_info

                chars = page.chars or []
                images = page.images or []
                content_boxes = _content_image_boxes(page)
                content_boxes_by_page[page_num] = content_boxes
                page_info["char_count"] = len(chars)
                page_info["embedded_image_count"] = len(images)
                page_info["content_image_count"] = len(content_boxes)

                title, bullets = _page_text_hints(page)
                hints = []
                if title:
                    hints.append(title)
                hints.extend(bullets)
                page_info["context_hints"] = hints[:6]

                tables = page.extract_tables() or []
                good = 0
                junk = 0
                for table_idx, table in enumerate(tables):
                    if not table:
                        continue
                    md = _table_to_markdown(table)
                    if not md.strip():
                        continue
                    is_junk = _is_junk_table(table)
                    if is_junk:
                        junk += 1
                    else:
                        good += 1
                    result["tables"].append(
                        {
                            "page_number": page_num,
                            "table_index": table_idx,
                            "markdown": md,
                            "empty_cell_ratio": round(_empty_cell_ratio(table), 4),
                            "row_count": len(table),
                            "is_junk": is_junk,
                        }
                    )
                page_info["good_table_count"] = good
                page_info["junk_table_count"] = junk
                page_info["likely_image_table"] = (
                    page_info["content_image_count"] >= 1 and good == 0
                )
                # Only claim "image table" captions when the page is flagged as such
                if page_info["likely_image_table"]:
                    page_info["image_caption"] = _build_image_caption(
                        page_num, title, bullets
                    )
                elif title:
                    short = title.replace("\n", " ").strip()
                    if len(short) > 80:
                        short = short[:80] + "…"
                    page_info["image_caption"] = f"第 {page_num} 頁截圖。標題「{short}」。"
                else:
                    page_info["image_caption"] = f"第 {page_num} 頁截圖。"

        result["tables_extracted_total"] = len(result["tables"])
        result["tables_extracted_good"] = sum(
            1 for t in result["tables"] if not t.get("is_junk")
        )
        # Keep legacy field as total for backward compatibility
        result["tables_extracted"] = result["tables_extracted_total"]
    except ImportError:
        result["errors"].append("pdfplumber 未安裝（pip install pdfplumber），略過表格擷取")
    except Exception as exc:
        result["errors"].append(f"pdfplumber 表格擷取失敗：{exc}")

    # --- Optional OCR for likely image-table pages ---
    # OCR gets a temporary render cropped to the embedded table image. The normal
    # 150-DPI full-page image remains the Markdown visual fallback.
    ocr_doc = None
    try:
        import fitz

        ocr_doc = fitz.open(str(input_pdf))
    except Exception:
        ocr_doc = None

    try:
        for page_info in result["pages"]:
            if not page_info.get("likely_image_table"):
                page_info["ocr_status"] = "not_attempted"
                continue
            img_rel = page_info.get("image_path") or ""
            img_path = output_dir / img_rel
            if not img_path.is_file():
                page_info["ocr_status"] = "skipped_no_image"
                continue

            ocr_path = img_path
            temp_path = images_dir / f".ocr-{page_info['page_number']}.png"
            try:
                if ocr_doc is not None:
                    page = ocr_doc.load_page(page_info["page_number"] - 1)
                    boxes = content_boxes_by_page.get(page_info["page_number"], [])
                    clip = None
                    if boxes:
                        largest = max(
                            boxes,
                            key=lambda box: (box[2] - box[0]) * (box[3] - box[1]),
                        )
                        clip = fitz.Rect(*largest)
                    page.get_pixmap(dpi=OCR_TABLE_DPI, clip=clip).save(str(temp_path))
                    ocr_path = temp_path
                preview, status = _try_ocr_text(ocr_path)
                page_info["ocr_preview"] = preview
                page_info["ocr_status"] = status
            finally:
                if temp_path.is_file():
                    temp_path.unlink()
    finally:
        if ocr_doc is not None:
            ocr_doc.close()

    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
