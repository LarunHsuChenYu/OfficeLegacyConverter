using System.Text;

namespace OfficeLegacyConverter;

internal static class XlsxMarkdownPipeline
{
    public static void ConvertFile(
        string inputPath,
        string outputDir,
        bool showOffice,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        var (resolvedInput, tempCleanup) = LegacyOfficeUpgrade.ResolveIfNeeded(
            inputPath, showOffice, progress, current, total, cancellationToken);

        try
        {
            var baseName = Path.GetFileNameWithoutExtension(inputPath);
            var rawMdPath = Path.Combine(outputDir, baseName + ".raw.md");
            var outputMdPath = Path.Combine(outputDir, baseName + ".md");

            progress.Report((current, total, $"Open XML 轉換 {resolvedInput} → {rawMdPath}..."));
            var xlsxConversion = XlsxMarkdownConverter.Convert(
                resolvedInput,
                sheetLog => progress.Report((current, total, $"  {sheetLog}")),
                cancellationToken);
            File.WriteAllText(rawMdPath, xlsxConversion.Markdown, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            progress.Report((current, total, $"  共轉換 {xlsxConversion.SheetCount} 個工作表"));
            foreach (var logEntry in xlsxConversion.ConversionLog)
                progress.Report((current, total, $"  {logEntry}"));

            progress.Report((current, total, "後處理：頁碼清理與 YAML 注入..."));
            var postResult = MarkdownPostProcessor.Process(
                rawMdPath,
                inputPath,
                outputDir,
                xlsxConversion: xlsxConversion,
                conversionEngine: "OpenXml");
            MarkdownPipeline.ReportPostProcess(progress, current, total, postResult, outputMdPath);
        }
        finally
        {
            tempCleanup?.Invoke();
        }
    }
}
