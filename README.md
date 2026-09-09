# OfficeLegacyConverter

Windows 桌面工具，用於：

- 舊版 Word、Excel、PowerPoint 格式升級
- Office 轉 PDF
- PDF 合併
- Office／PDF 轉 Markdown
- Outlook `.msg` 擷取為 Markdown 與附件

## 版本

目前版本：`1.1.2`

發布日期：`2026-09-09`

版本與日期會顯示在主畫面抬頭及 Windows 視窗標題。

## PDF 影像表格

PDF 增強流程會：

1. 用 pdfplumber 擷取文字層表格。
2. 排除全頁背景與小型 Logo，只把內容區的大型嵌入圖片視為影像表候選。
3. 對缺表頁裁切表格區，以 400 DPI 執行 PSM 3／6，再依通用表格數值訊號選出單一 OCR 結果。
4. 在 `_quality.json` 區分 `tables_extracted_total` 與 `tables_extracted_good`。

若影像表格尚未還原為 Markdown 結構化表格，`semantic_coverage` 會列出缺表頁數，並區分是否已有 OCR 可搜尋文字，不再宣稱固定的 `90%+`。

`missing_table` 代表結構化表格仍缺失；`recovery_status` 另行標示已有未經逐格驗證的 OCR，或仍只能查看圖片。`.raw.md` 保留原始擷取內容，OCR 與缺表標記只寫入強化後的 `.md`。

## OCR

原始碼執行需要 Tesseract 5。Windows 可用：

```powershell
winget install --id UB-Mannheim.TesseractOCR --exact
```

繁中與英文語言資料放在 `scripts/tessdata/`。正式 Release 壓縮檔包含 Tesseract 執行檔，不需另外安裝。

## Mail analysis 交接

`.msg` 轉換只負責證據擷取，不等於郵件分析完成。

- 輸出選擇 `{mail-root}/extracted/` 時，原始 `.msg` 會保留到同層 `raw/`。
- 附件位於 `extracted/attachments/{mail-id}/`。
- 產出標示 `analysis_status: extraction_only`。
- 後續仍須完成最上層正文 claim ledger、完整 thread 與附件內容核對。

## 建置

```powershell
dotnet build OfficeLegacyConverter.csproj -c Release
dotnet publish OfficeLegacyConverter.csproj -c Release -r win-x64 --self-contained true
```

需求：Windows、.NET 10 SDK。從原始碼執行 PDF／PPTX Markdown 轉換時，Python 3.10+ 需安裝 `markitdown`、`pdfplumber` 與 `pymupdf`。
