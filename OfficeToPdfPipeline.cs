namespace OfficeLegacyConverter;

internal static class OfficeToPdfPipeline
{
    public static readonly string[] SupportedExtensions =
        [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"];

    private static readonly HashSet<string> WordExtensions = new(StringComparer.OrdinalIgnoreCase) { ".doc", ".docx" };
    private static readonly HashSet<string> ExcelExtensions = new(StringComparer.OrdinalIgnoreCase) { ".xls", ".xlsx" };
    private static readonly HashSet<string> PowerPointExtensions = new(StringComparer.OrdinalIgnoreCase) { ".ppt", ".pptx" };

    public static BatchResult ConvertFiles(
        IReadOnlyList<string> paths,
        bool replace,
        bool showOffice,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        var total = paths.Count;
        var result = new BatchResult { Total = total };
        for (var i = 0; i < total; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var inputPath = paths[i];
            var current = i + 1;

            try
            {
                ConvertFile(inputPath, showOffice, current, total, progress, cancellationToken);

                if (replace)
                {
                    File.Delete(inputPath);
                    progress.Report((current, total, $"已刪除原始檔：{inputPath}"));
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                result.Failures.Add($"{Path.GetFileName(inputPath)}：{ex.Message}");
                progress.Report((current, total, $"錯誤 - {inputPath}：{ex.Message}"));
            }
        }

        return result;
    }

    private static void ConvertFile(
        string inputPath,
        bool showOffice,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        var ext = Path.GetExtension(inputPath);
        var pdfPath = Path.ChangeExtension(inputPath, ".pdf");
        var (resolvedInput, tempCleanup) = LegacyOfficeUpgrade.ResolveIfNeeded(
            inputPath, showOffice, progress, current, total, cancellationToken);

        try
        {
            if (WordExtensions.Contains(ext))
            {
                WordPdfConverter.ConvertFile(resolvedInput, pdfPath, showOffice, current, total, progress, cancellationToken);
                return;
            }

            if (ExcelExtensions.Contains(ext))
            {
                ExcelPdfConverter.ConvertFile(resolvedInput, pdfPath, showOffice, current, total, progress, cancellationToken);
                return;
            }

            if (PowerPointExtensions.Contains(ext))
            {
                PowerPointPdfConverter.ConvertFile(resolvedInput, pdfPath, showOffice, current, total, progress, cancellationToken);
                return;
            }

            throw new NotSupportedException($"不支援的副檔名：{ext}");
        }
        finally
        {
            tempCleanup?.Invoke();
        }
    }
}
