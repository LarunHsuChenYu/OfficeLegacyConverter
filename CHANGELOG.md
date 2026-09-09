# Changelog

## 1.1.0 - 2026-09-09

- 修正純文字頁被誤判為 `missing_table`。
- 影像表偵測排除全頁背景與小型 Logo。
- PDF 影像表採 300 DPI Tesseract OCR，繁中與英文文字可寫入 Markdown。
- 語意覆蓋率改為依缺表頁數揭露，不再顯示固定 `90%+`。
- 主畫面抬頭及視窗標題顯示版本與發布日期。
- `.msg` 輸出標記為 `extraction_only`；使用 `Mail/extracted/` 時保留 Primary `.msg` 到 `Mail/raw/`。
