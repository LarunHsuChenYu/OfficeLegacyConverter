using System.Runtime.InteropServices;
using System.Text;

namespace OfficeLegacyConverter;

internal static class PowerPointComHelper
{
    private const int PpAlertsNone = 1;
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;
    private const int PpWindowMinimized = 2;

    /// <summary>
    /// 建立獨立 PowerPoint 執行個體供 PDF 匯出，避免附加至使用者已開啟的 PowerPoint。
    /// </summary>
    public static dynamic CreateDedicatedPowerPointForPdfExport()
    {
        var powerPointType = Type.GetTypeFromProgID("PowerPoint.Application")
            ?? throw new InvalidOperationException("找不到 Microsoft PowerPoint，請確認已安裝。");

        return Activator.CreateInstance(powerPointType)!;
    }

    /// <summary>
    /// 設定 PowerPoint 應用程式可見性。部分版本/政策不允許隱藏視窗，失敗時改為顯示或最小化。
    /// </summary>
    public static void TrySetApplicationVisible(dynamic powerPoint, bool show)
    {
        if (show)
        {
            try { powerPoint.Visible = MsoTrue; } catch (COMException) { /* ignore */ }
            return;
        }

        try
        {
            powerPoint.Visible = MsoFalse;
        }
        catch (COMException)
        {
            // 部分環境不允許隱藏 PowerPoint 視窗；改為顯示並嘗試最小化。
            try { powerPoint.Visible = MsoTrue; } catch (COMException) { /* ignore */ }
            try { powerPoint.WindowState = PpWindowMinimized; } catch (COMException) { /* ignore */ }
        }
    }

    public static void ConfigurePowerPointForExport(dynamic powerPoint)
    {
        try { powerPoint.DisplayAlerts = PpAlertsNone; } catch (COMException) { /* ignore */ }
    }

    public static InvalidOperationException WrapExportException(string step, Exception ex)
    {
        var message = new StringBuilder();
        message.Append($"PowerPoint PDF 匯出失敗（步驟：{step}）。");

        if (ex is COMException com)
            message.Append($" HRESULT: 0x{com.HResult:X8}。");

        message.Append($" {ex.Message}");

        if (ex.InnerException is not null)
            message.Append($" InnerException: {ex.InnerException.Message}");

        return new InvalidOperationException(message.ToString(), ex);
    }
}
