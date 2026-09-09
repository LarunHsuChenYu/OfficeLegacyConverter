using System.Runtime.InteropServices;

namespace OfficeLegacyConverter;

internal static class WordConverter
{
    private const int WdFormatXmlDocument = 12;

    public static void ConvertFile(
        string inputPath,
        string outputPath,
        bool showWord,
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
            word.Documents.Open(inputPath);
            word.ActiveDocument.SaveAs2(outputPath, WdFormatXmlDocument);
            word.ActiveDocument.Close();
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

    public static BatchResult ConvertFiles(
        IReadOnlyList<string> paths,
        bool replace,
        bool showWord,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        dynamic? word = null;
        var total = paths.Count;
        var result = new BatchResult { Total = total };
        try
        {
            var wordType = Type.GetTypeFromProgID("Word.Application")
                ?? throw new InvalidOperationException("找不到 Microsoft Word，請確認已安裝。");

            word = Activator.CreateInstance(wordType)!;
            WordComHelper.TrySetApplicationVisible(word, showWord);
            WordComHelper.ConfigureWordForExport(word);

            for (var i = 0; i < total; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var path = paths[i];
                var current = i + 1;
                progress.Report((current, total, $"開啟 {path}..."));

                try
                {
                    word.Documents.Open(path);
                    var savePath = Path.ChangeExtension(path, ".docx");
                    word.ActiveDocument.SaveAs2(savePath, WdFormatXmlDocument);
                    word.ActiveDocument.Close();

                    if (replace)
                        File.Delete(path);

                    progress.Report((current, total, $"完成：{Path.GetFileName(path)} → {Path.GetFileName(savePath)}"));
                }
                catch (Exception ex)
                {
                    try { word.ActiveDocument?.Close(false); } catch { /* ignore */ }
                    result.Failures.Add($"{Path.GetFileName(path)}：{ex.Message}");
                    progress.Report((current, total, $"錯誤 - {path}：{ex.Message}"));
                }
            }
        }
        finally
        {
            if (word is not null)
            {
                try { word.Quit(); } catch { /* ignore */ }
                Marshal.ReleaseComObject(word);
            }
        }

        return result;
    }
}
