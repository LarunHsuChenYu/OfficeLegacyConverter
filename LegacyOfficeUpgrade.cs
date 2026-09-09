namespace OfficeLegacyConverter;

/// <summary>
/// 將 .doc / .xls / .ppt 升級為對應新格式暫存檔，供 PDF / Markdown 管線共用。
/// </summary>
internal static class LegacyOfficeUpgrade
{
    public static (string ResolvedInput, Action? Cleanup) ResolveIfNeeded(
        string inputPath,
        bool showOffice,
        IProgress<(int current, int total, string message)> progress,
        int current,
        int total,
        CancellationToken cancellationToken)
    {
        var ext = Path.GetExtension(inputPath);
        var modernExt = ext.ToLowerInvariant() switch
        {
            ".doc" => ".docx",
            ".xls" => ".xlsx",
            ".ppt" => ".pptx",
            _ => (string?)null
        };

        if (modernExt is null)
            return (inputPath, null);

        var tempDir = Path.Combine(Path.GetTempPath(), "OfficeLegacyConverter", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var upgradedPath = Path.Combine(tempDir, Path.GetFileName(Path.ChangeExtension(inputPath, modernExt)));
        progress.Report((current, total, $"升級 {inputPath} → {upgradedPath}..."));

        switch (modernExt)
        {
            case ".docx":
                WordConverter.ConvertFile(inputPath, upgradedPath, showOffice, cancellationToken);
                break;
            case ".xlsx":
                ExcelConverter.ConvertFile(inputPath, upgradedPath, showOffice, cancellationToken);
                break;
            case ".pptx":
                PowerPointConverter.ConvertFile(inputPath, upgradedPath, showOffice, cancellationToken);
                break;
        }

        return (upgradedPath, () => CleanupTemp(tempDir, upgradedPath));
    }

    private static void CleanupTemp(string tempDir, string upgradedPath)
    {
        try
        {
            if (File.Exists(upgradedPath))
                File.Delete(upgradedPath);
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, recursive: true);
        }
        catch { /* ignore */ }
    }
}
