using System.Runtime.InteropServices;

namespace OfficeLegacyConverter;

internal static class PowerPointConverter
{
    private const int PpSaveAsOpenXmlPresentation = 24;

    public static void ConvertFile(
        string inputPath,
        string outputPath,
        bool showPowerPoint,
        CancellationToken cancellationToken)
    {
        dynamic? powerPoint = null;
        try
        {
            powerPoint = PowerPointComHelper.CreateDedicatedPowerPointForPdfExport();
            PowerPointComHelper.TrySetApplicationVisible(powerPoint, showPowerPoint);
            PowerPointComHelper.ConfigurePowerPointForExport(powerPoint);

            cancellationToken.ThrowIfCancellationRequested();
            powerPoint.Presentations.Open(inputPath);
            powerPoint.ActivePresentation.SaveAs(outputPath, PpSaveAsOpenXmlPresentation);
            powerPoint.ActivePresentation.Close();
        }
        finally
        {
            if (powerPoint is not null)
            {
                try { powerPoint.Quit(); } catch { /* ignore */ }
                Marshal.ReleaseComObject(powerPoint);
            }
        }
    }

    public static BatchResult ConvertFiles(
        IReadOnlyList<string> paths,
        bool replace,
        bool showPowerPoint,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        dynamic? powerPoint = null;
        var total = paths.Count;
        var result = new BatchResult { Total = total };
        try
        {
            powerPoint = PowerPointComHelper.CreateDedicatedPowerPointForPdfExport();
            PowerPointComHelper.TrySetApplicationVisible(powerPoint, showPowerPoint);
            PowerPointComHelper.ConfigurePowerPointForExport(powerPoint);

            for (var i = 0; i < total; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var path = paths[i];
                var current = i + 1;
                progress.Report((current, total, $"開啟 {path}..."));

                try
                {
                    powerPoint.Presentations.Open(path);
                    var savePath = Path.ChangeExtension(path, ".pptx");
                    powerPoint.ActivePresentation.SaveAs(savePath, PpSaveAsOpenXmlPresentation);
                    powerPoint.ActivePresentation.Close();

                    if (replace)
                        File.Delete(path);

                    progress.Report((current, total, $"完成：{Path.GetFileName(path)} → {Path.GetFileName(savePath)}"));
                }
                catch (Exception ex)
                {
                    try { powerPoint.ActivePresentation?.Close(); } catch { /* ignore */ }
                    var message = ex is COMException com
                        ? $"{ex.Message} (HRESULT: 0x{com.HResult:X8})"
                        : ex.Message;
                    result.Failures.Add($"{Path.GetFileName(path)}：{message}");
                    progress.Report((current, total, $"錯誤 - {path}：{message}"));
                }
            }
        }
        finally
        {
            if (powerPoint is not null)
            {
                try { powerPoint.Quit(); } catch { /* ignore */ }
                Marshal.ReleaseComObject(powerPoint);
            }
        }

        return result;
    }
}
