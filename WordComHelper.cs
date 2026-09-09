using System.Runtime.InteropServices;
using System.Text;

namespace OfficeLegacyConverter;

internal static class WordComHelper
{
    private const int WdAlertsNone = 0;
    private const int WdWindowStateMinimize = 2;

    /// <summary>
    /// 設定 Word 應用程式可見性。部分環境不允許隱藏視窗，失敗時改為顯示或最小化。
    /// </summary>
    public static void TrySetApplicationVisible(dynamic word, bool show)
    {
        if (show)
        {
            try { word.Visible = true; } catch (COMException) { /* ignore */ }
            return;
        }

        try
        {
            word.Visible = false;
        }
        catch (COMException)
        {
            try { word.Visible = true; } catch (COMException) { /* ignore */ }
            try
            {
                if (word.ActiveWindow is not null)
                    word.ActiveWindow.WindowState = WdWindowStateMinimize;
            }
            catch (COMException) { /* ignore */ }
        }
    }

    public static void ConfigureWordForExport(dynamic word)
    {
        try { word.DisplayAlerts = WdAlertsNone; } catch (COMException) { /* ignore */ }
    }

    public static InvalidOperationException WrapExportException(string step, Exception ex)
    {
        var message = new StringBuilder();
        message.Append($"Word PDF 匯出失敗（步驟：{step}）。");

        if (ex is COMException com)
            message.Append($" HRESULT: 0x{com.HResult:X8}。");

        message.Append($" {ex.Message}");

        if (ex.InnerException is not null)
            message.Append($" InnerException: {ex.InnerException.Message}");

        return new InvalidOperationException(message.ToString(), ex);
    }
}
