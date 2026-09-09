using System.Runtime.InteropServices;

namespace OfficeLegacyConverter;

internal static class ExcelConverter
{
    private const int XlOpenXmlWorkbook = 51;

    public static void ConvertFile(
        string inputPath,
        string outputPath,
        bool showExcel,
        CancellationToken cancellationToken)
    {
        dynamic? excel = null;
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("找不到 Microsoft Excel，請確認已安裝。");

            excel = Activator.CreateInstance(excelType)!;
            excel.Visible = showExcel;
            excel.DisplayAlerts = false;

            cancellationToken.ThrowIfCancellationRequested();
            excel.Workbooks.Open(inputPath);
            excel.ActiveWorkbook.SaveAs(outputPath, XlOpenXmlWorkbook);
            excel.ActiveWorkbook.Close();
        }
        finally
        {
            if (excel is not null)
            {
                try { excel.Quit(); } catch { /* ignore */ }
                Marshal.ReleaseComObject(excel);
            }
        }
    }

    public static BatchResult ConvertFiles(
        IReadOnlyList<string> paths,
        bool replace,
        bool showExcel,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        dynamic? excel = null;
        var total = paths.Count;
        var result = new BatchResult { Total = total };
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("找不到 Microsoft Excel，請確認已安裝。");

            excel = Activator.CreateInstance(excelType)!;
            excel.Visible = showExcel;
            excel.DisplayAlerts = false;

            for (var i = 0; i < total; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var path = paths[i];
                var current = i + 1;
                progress.Report((current, total, $"開啟 {path}..."));

                try
                {
                    excel.Workbooks.Open(path);
                    var savePath = Path.ChangeExtension(path, ".xlsx");
                    excel.ActiveWorkbook.SaveAs(savePath, XlOpenXmlWorkbook);
                    excel.ActiveWorkbook.Close();

                    if (replace)
                        File.Delete(path);

                    progress.Report((current, total, $"完成：{Path.GetFileName(path)} → {Path.GetFileName(savePath)}"));
                }
                catch (Exception ex)
                {
                    try { excel.ActiveWorkbook?.Close(false); } catch { /* ignore */ }
                    result.Failures.Add($"{Path.GetFileName(path)}：{ex.Message}");
                    progress.Report((current, total, $"錯誤 - {path}：{ex.Message}"));
                }
            }
        }
        finally
        {
            if (excel is not null)
            {
                try { excel.Quit(); } catch { /* ignore */ }
                Marshal.ReleaseComObject(excel);
            }
        }

        return result;
    }
}
