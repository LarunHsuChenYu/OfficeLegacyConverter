# Changelog

## 1.1.1 - 2026-09-09

- 移除 OCR 1500 字元與 40 行限制，完整保留影像表文字。
- OCR 改為裁切表格區、400 DPI，搭配 PSM 3 與 PSM 11 補充辨識。
- 修復 p.4、p.7、p.10 尾端數值截斷，以及 p.9 RI Sales 的 `13`、`2`。
- `semantic_coverage` 改為區分「未還原結構化表格」與「已有 OCR 可搜尋文字」。
- `_quality.json` 新增 `ocr_text_length`、`ocr_truncated`，並保留舊欄位相容性。

## 1.1.0 - 2026-09-09

- 修正純文字頁被誤判為 `missing_table`。
- 影像表偵測排除全頁背景與小型 Logo。
- PDF 影像表採 300 DPI Tesseract OCR，繁中與英文文字可寫入 Markdown。
- 語意覆蓋率改為依缺表頁數揭露，不再顯示固定 `90%+`。
- 主畫面抬頭及視窗標題顯示版本與發布日期。
- `.msg` 輸出標記為 `extraction_only`；使用 `Mail/extracted/` 時保留 Primary `.msg` 到 `Mail/raw/`。
- 升級 MsgReader 6.1.1 與 PDFsharp 6.2.4，移除 OpenMcdf 3.0.0 的已知中度弱點。
