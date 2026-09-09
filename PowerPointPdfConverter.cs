using System.Runtime.InteropServices;

namespace OfficeLegacyConverter;

internal static class PowerPointPdfConverter
{
    private const int PpSaveAsPdf = 32;

    public static void ConvertFile(
        string inputPath,
        string pdfPath,
        bool showPowerPoint,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        dynamic? powerPoint = null;
        try
        {
            powerPoint = PowerPointComHelper.CreateDedicatedPowerPointForPdfExport();
            PowerPointComHelper.TrySetApplicationVisible(powerPoint, showPowerPoint);
            PowerPointComHelper.ConfigurePowerPointForExport(powerPoint);

            cancellationToken.ThrowIfCancellationRequested();
            progress.Report((current, total, $"開啟 {inputPath}..."));

            powerPoint.Presentations.Open(inputPath);
            powerPoint.ActivePresentation.SaveAs(pdfPath, PpSaveAsPdf);
            powerPoint.ActivePresentation.Close();

            progress.Report((current, total, $"完成：{Path.GetFileName(inputPath)} → {Path.GetFileName(pdfPath)}"));
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw PowerPointComHelper.WrapExportException("PowerPoint 轉 PDF", ex);
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
}
