namespace OfficeLegacyConverter;

internal static class MarkdownPipeline
{
    public static readonly string[] SupportedExtensions =
        [".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"];

    public static void ConvertFiles(
        IReadOnlyList<string> paths,
        string outputFolder,
        bool deleteOriginal,
        bool showOffice,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        if (paths.Any(RequiresMarkItDown) && !MarkItDownConverter.IsEnvironmentReady)
            throw new InvalidOperationException(MarkItDownConverter.EnvironmentStatus);

        var outputDir = Path.GetFullPath(outputFolder);
        if (!Directory.Exists(outputDir))
            throw new DirectoryNotFoundException($"輸出資料夾不存在：{outputDir}");

        var total = paths.Count;
        for (var i = 0; i < total; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var inputPath = paths[i];
            var current = i + 1;
            var ext = Path.GetExtension(inputPath);

            try
            {
                RouteConversion(inputPath, outputDir, showOffice, current, total, progress, ext, cancellationToken);

                if (deleteOriginal)
                {
                    File.Delete(inputPath);
                    progress.Report((current, total, $"已刪除原始檔：{inputPath}"));
                }
            }
            catch (Exception ex)
            {
                progress.Report((current, total, $"錯誤 - {inputPath}：{ex.Message}"));
            }
        }
    }

    private static void RouteConversion(
        string inputPath,
        string outputDir,
        bool showOffice,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        string ext,
        CancellationToken cancellationToken)
    {
        if (ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            PdfMarkdownPipeline.ConvertFile(inputPath, outputDir, current, total, progress, cancellationToken);
            return;
        }

        if (IsDocxInput(ext))
        {
            DocxMarkdownPipeline.ConvertFile(inputPath, outputDir, showOffice, current, total, progress, cancellationToken);
            return;
        }

        if (IsXlsxInput(ext))
        {
            XlsxMarkdownPipeline.ConvertFile(inputPath, outputDir, showOffice, current, total, progress, cancellationToken);
            return;
        }

        if (IsPptxInput(ext))
        {
            PptxMarkdownPipeline.ConvertFile(inputPath, outputDir, showOffice, current, total, progress, cancellationToken);
            return;
        }

        throw new NotSupportedException($"不支援的 Markdown 轉換格式：{ext}");
    }

    internal static void ReportPostProcess(
        IProgress<(int current, int total, string message)> progress,
        int current,
        int total,
        PostProcessResult postResult,
        string outputMdPath)
    {
        foreach (var entry in postResult.CleanupLog.Take(8))
            progress.Report((current, total, $"  清理：{entry}"));
        if (postResult.CleanupLog.Count > 8)
            progress.Report((current, total, $"  …共 {postResult.CleanupLog.Count} 項清理"));

        if (postResult.DetectedGaps.Count > 0)
            progress.Report((current, total, $"  偵測到 {postResult.DetectedGaps.Count} 個已知缺口（見 _quality.json）"));

        progress.Report((current, total, $"完成：{outputMdPath}（已套用 YAML 後設資料與頁碼清理）"));
    }

    private static bool IsDocxInput(string ext) =>
        ext.Equals(".docx", StringComparison.OrdinalIgnoreCase)
        || ext.Equals(".doc", StringComparison.OrdinalIgnoreCase);

    private static bool IsXlsxInput(string ext) =>
        ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
        || ext.Equals(".xls", StringComparison.OrdinalIgnoreCase);

    private static bool IsPptxInput(string ext) =>
        ext.Equals(".pptx", StringComparison.OrdinalIgnoreCase)
        || ext.Equals(".ppt", StringComparison.OrdinalIgnoreCase);

    private static bool RequiresMarkItDown(string path)
    {
        var ext = Path.GetExtension(path);
        return ext.Equals(".pdf", StringComparison.OrdinalIgnoreCase) || IsPptxInput(ext);
    }
}
