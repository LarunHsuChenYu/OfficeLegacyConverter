using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace OfficeLegacyConverter;

internal static class PdfMerger
{
    public static void MergeFiles(
        IReadOnlyList<string> inputPaths,
        string outputPath,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        outputPath = Path.GetFullPath(outputPath);
        using var outputDocument = new PdfDocument();
        var total = inputPaths.Count;

        for (var i = 0; i < total; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var path = inputPaths[i];
            var current = i + 1;
            progress.Report((current, total, $"處理 {path}..."));

            try
            {
                using var inputDocument = PdfReader.Open(path, PdfDocumentOpenMode.Import);
                var pageCount = inputDocument.PageCount;
                for (var pageIndex = 0; pageIndex < pageCount; pageIndex++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    outputDocument.AddPage(inputDocument.Pages[pageIndex]);
                }

                progress.Report((current, total, $"已加入 {pageCount} 頁：{Path.GetFileName(path)}"));
            }
            catch (Exception ex)
            {
                progress.Report((current, total, $"錯誤 - {path}：{ex.Message}"));
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        progress.Report((total, total, $"儲存 {outputPath}..."));
        outputDocument.Save(outputPath);
        progress.Report((total, total, $"完成：{outputPath}"));
    }

    public static void MergeFilesOrThrow(IReadOnlyList<string> inputPaths, string outputPath)
    {
        outputPath = Path.GetFullPath(outputPath);
        using var outputDocument = new PdfDocument();

        foreach (var path in inputPaths)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"合併 PDF 時找不到檔案：{path}");

            using var inputDocument = PdfReader.Open(path, PdfDocumentOpenMode.Import);
            for (var pageIndex = 0; pageIndex < inputDocument.PageCount; pageIndex++)
                outputDocument.AddPage(inputDocument.Pages[pageIndex]);
        }

        outputDocument.Save(outputPath);
    }
}
