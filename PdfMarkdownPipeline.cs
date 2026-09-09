namespace OfficeLegacyConverter;

internal static class PdfMarkdownPipeline
{
    public static void ConvertFile(
        string inputPath,
        string outputDir,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        var baseName = Path.GetFileNameWithoutExtension(inputPath);
        var rawMdPath = Path.Combine(outputDir, baseName + ".raw.md");
        var outputMdPath = Path.Combine(outputDir, baseName + ".md");

        progress.Report((current, total, $"轉換 {inputPath} → {rawMdPath}..."));
        MarkItDownConverter.ConvertFile(inputPath, rawMdPath, cancellationToken);

        progress.Report((current, total, "PDF 增強：擷取表格與頁面圖片..."));
        var pdfEnhancement = PdfEnhancementService.Enhance(inputPath, outputDir, baseName, cancellationToken);
        progress.Report((current, total,
            $"  已擷取 {pdfEnhancement.PagesExported} 頁圖片、表格 total={pdfEnhancement.TablesExtractedTotal} / good={pdfEnhancement.TablesExtractedGood}"));
        var ocrPages = pdfEnhancement.Pages.Count(p => p.OcrStatus == "ok");
        if (ocrPages > 0)
            progress.Report((current, total, $"  OCR：已產生 {ocrPages} 頁可搜尋文字摘要"));
        if (pdfEnhancement.Pages.Any(p => p.OcrStatus is "skipped_no_executable" or "skipped_no_engine"))
            progress.Report((current, total, "  OCR：找不到 Tesseract；可執行 winget install UB-Mannheim.TesseractOCR"));
        if (pdfEnhancement.Pages.Any(p => p.OcrStatus == "skipped_no_language"))
            progress.Report((current, total, "  OCR：缺少 chi_tra／eng 語言資料，請確認 scripts/tessdata"));
        foreach (var err in pdfEnhancement.Errors.Take(3))
            progress.Report((current, total, $"  PDF 增強：{err}"));
        if (!string.IsNullOrEmpty(pdfEnhancement.Error))
            progress.Report((current, total, $"  PDF 增強警告：{pdfEnhancement.Error}"));

        progress.Report((current, total, "後處理：頁碼清理與 YAML 注入..."));
        var postResult = MarkdownPostProcessor.Process(
            rawMdPath,
            inputPath,
            outputDir,
            pdfEnhancement,
            conversionEngine: "MarkItDown");
        MarkdownPipeline.ReportPostProcess(progress, current, total, postResult, outputMdPath);
    }
}
