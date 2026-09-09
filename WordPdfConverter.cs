using System.Runtime.InteropServices;

namespace OfficeLegacyConverter;

internal static class WordPdfConverter
{
    private const int WdFormatPdf = 17;

    public static void ConvertFile(
        string inputPath,
        string pdfPath,
        bool showWord,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        dynamic? word = null;
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("找不到 Microsoft Word，請確認已安裝。");

            word = Activator.CreateInstance(wordType)!;
            WordComHelper.TrySetApplicationVisible(word, showWord);
            WordComHelper.ConfigureWordForExport(word);

            cancellationToken.ThrowIfCancellationRequested();
            progress.Report((current, total, $"開啟 {inputPath}..."));

            word.Documents.Open(inputPath);
            word.ActiveDocument.SaveAs2(pdfPath, WdFormatPdf);
            word.ActiveDocument.Close();

            progress.Report((current, total, $"完成：{Path.GetFileName(inputPath)} → {Path.GetFileName(pdfPath)}"));
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw WordComHelper.WrapExportException("Word 轉 PDF", ex);
        }
        finally
        {
            if (word is not null)
            {
                try { word.Quit(); } catch { /* ignore */ }
                Marshal.ReleaseComObject(word);
            }
        }
    }
}
