using System.Text.RegularExpressions;

namespace OfficeLegacyConverter;

/// <summary>
/// 非 modal 轉換進度視窗；以 Show() 顯示，不阻塞 UI 或 STA 背景執行緒。
/// </summary>
internal sealed class ConversionProgressForm : Form
{
    private static readonly Regex SubProgressRegex = new(
        @"[（(](\d+)/(\d+)[）)]|(\d+)/(\d+)",
        RegexOptions.Compiled);

    private readonly Label _lblMessage;
    private readonly Label _lblFileProgress;
    private readonly ProgressBar _progressBar;

    public ConversionProgressForm()
    {
        Text = "轉換進度";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        TopMost = true;
        ShowInTaskbar = false;
        ClientSize = new Size(420, 130);

        _lblMessage = new Label
        {
            AutoSize = false,
            Location = new Point(16, 16),
            Size = new Size(388, 48),
            Font = new Font(Font.FontFamily, 11f, FontStyle.Bold),
            Text = "準備中…",
        };

        _lblFileProgress = new Label
        {
            AutoSize = true,
            Location = new Point(16, 72),
            Text = "",
        };

        _progressBar = new ProgressBar
        {
            Location = new Point(16, 96),
            Size = new Size(388, 23),
            Style = ProgressBarStyle.Continuous,
        };

        Controls.Add(_lblMessage);
        Controls.Add(_lblFileProgress);
        Controls.Add(_progressBar);
    }

    public void UpdateProgress(int current, int total, string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(UpdateProgress, current, total, message);
            return;
        }

        _lblMessage.Text = message;

        if (total > 0 && current > 0)
            _lblFileProgress.Text = $"檔案 {current} / {total}";
        else if (total > 0)
            _lblFileProgress.Text = $"共 {total} 個檔案";
        else
            _lblFileProgress.Text = "";

        if (TryParseSubProgress(message, out var subCurrent, out var subTotal))
        {
            _progressBar.Maximum = subTotal;
            _progressBar.Value = Math.Clamp(subCurrent, 0, subTotal);
        }
        else if (total > 0)
        {
            _progressBar.Maximum = total;
            _progressBar.Value = Math.Clamp(current, 0, total);
        }
    }

    public void CloseSafely()
    {
        if (IsDisposed)
            return;

        if (InvokeRequired)
        {
            BeginInvoke(CloseSafely);
            return;
        }

        Close();
    }

    private static bool TryParseSubProgress(string message, out int subCurrent, out int subTotal)
    {
        subCurrent = 0;
        subTotal = 0;

        var match = SubProgressRegex.Match(message);
        if (!match.Success)
            return false;

        if (match.Groups[1].Success)
        {
            subCurrent = int.Parse(match.Groups[1].Value);
            subTotal = int.Parse(match.Groups[2].Value);
        }
        else
        {
            subCurrent = int.Parse(match.Groups[3].Value);
            subTotal = int.Parse(match.Groups[4].Value);
        }

        return subTotal > 0;
    }
}
