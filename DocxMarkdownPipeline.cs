using System.Text;

namespace OfficeLegacyConverter;

internal static class DocxMarkdownPipeline
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
            var docxConversion = DocxMarkdownConverter.Convert(resolvedInput, outputDir, baseName, cancellationToken);
            File.WriteAllText(rawMdPath, docxConversion.Markdown, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            if (docxConversion.Images.Count > 0)
            {
                progress.Report((current, total, $"  已匯出 {docxConversion.Images.Count} 張內嵌圖片"));
                foreach (var logEntry in docxConversion.ExportLog)
                    progress.Report((current, total, $"  圖片：{logEntry}"));
            }

            progress.Report((current, total, "後處理：頁碼清理與 YAML 注入..."));
            var postResult = MarkdownPostProcessor.Process(
                rawMdPath,
                inputPath,
                outputDir,
                docxConversion: docxConversion,
                conversionEngine: "OpenXml");
            MarkdownPipeline.ReportPostProcess(progress, current, total, postResult, outputMdPath);
        }
        finally
        {
            tempCleanup?.Invoke();
        }
    }
}
