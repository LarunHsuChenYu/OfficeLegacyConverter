using System.Runtime.InteropServices;

namespace OfficeLegacyConverter;

/// <summary>透過 Windows Shell 取得副檔名關聯圖示。</summary>
internal static class ShellIcons
{
    private const uint ShgfiIcon = 0x000000100;
    private const uint ShgfiSmallIcon = 0x000000001;
    private const uint ShgfiLargeIcon = 0x000000000;
    private const uint ShgfiUseFileAttributes = 0x000000010;
    private const uint FileAttributeNormal = 0x00000080;

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ShFileInfo
    {
        public IntPtr hIcon;
        public int iIcon;
        public uint dwAttributes;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
        public string szDisplayName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
        public string szTypeName;
    }

    [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr SHGetFileInfo(
        string pszPath,
        uint dwFileAttributes,
        ref ShFileInfo psfi,
        uint cbFileInfo,
        uint uFlags);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool DestroyIcon(IntPtr hIcon);

    private static readonly Dictionary<string, Icon> Cache = new(StringComparer.OrdinalIgnoreCase);

    public static Icon GetSmallIcon(string extension)
    {
        var key = NormalizeExtension(extension);
        if (Cache.TryGetValue(key, out var cached))
            return cached;

        var icon = ExtractIcon(key, large: false) ?? CreateFallbackIcon(key, 16);
        Cache[key] = icon;
        return icon;
    }

    public static Icon GetAppIcon()
    {
        // 以 Word 文件圖示作為應用程式圖示（較具辨識度）
        return ExtractIcon(".docx", large: true) ?? CreateFallbackIcon(".docx", 32);
    }

    /// <summary>Office → PDF 用：三色方塊代表 Word / Excel / PowerPoint。</summary>
    public static Icon GetOfficeSuiteIcon()
    {
        const string key = "__office_suite__";
        if (Cache.TryGetValue(key, out var cached))
            return cached;

        using var bmp = new Bitmap(16, 16);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Transparent);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.None;
            using (var b = new SolidBrush(UiTheme.Word))
                g.FillRectangle(b, 1, 1, 6, 6);
            using (var b = new SolidBrush(UiTheme.Excel))
                g.FillRectangle(b, 9, 1, 6, 6);
            using (var b = new SolidBrush(UiTheme.PowerPoint))
                g.FillRectangle(b, 1, 9, 6, 6);
            using (var b = new SolidBrush(UiTheme.Pdf))
                g.FillRectangle(b, 9, 9, 6, 6);
        }

        var icon = IconFromBitmap(bmp);
        Cache[key] = icon;
        return icon;
    }

    private static string NormalizeExtension(string extension)
    {
        if (string.IsNullOrWhiteSpace(extension))
            return ".txt";
        return extension.StartsWith('.') ? extension.ToLowerInvariant() : "." + extension.ToLowerInvariant();
    }

    private static Icon? ExtractIcon(string extension, bool large)
    {
        try
        {
            var shfi = new ShFileInfo();
            var flags = ShgfiIcon | ShgfiUseFileAttributes | (large ? ShgfiLargeIcon : ShgfiSmallIcon);
            var result = SHGetFileInfo(
                "dummy" + extension,
                FileAttributeNormal,
                ref shfi,
                (uint)Marshal.SizeOf<ShFileInfo>(),
                flags);

            if (result == IntPtr.Zero || shfi.hIcon == IntPtr.Zero)
                return null;

            using var temp = Icon.FromHandle(shfi.hIcon);
            var clone = (Icon)temp.Clone();
            DestroyIcon(shfi.hIcon);
            return clone;
        }
        catch
        {
            return null;
        }
    }

    private static Icon CreateFallbackIcon(string extension, int size)
    {
        var accent = extension switch
        {
            ".doc" or ".docx" => UiTheme.Word,
            ".xls" or ".xlsx" => UiTheme.Excel,
            ".ppt" or ".pptx" => UiTheme.PowerPoint,
            ".pdf" => UiTheme.Pdf,
            ".md" => UiTheme.Markdown,
            _ => UiTheme.OfficePdf
        };

        var label = extension.TrimStart('.').ToUpperInvariant();
        if (label.Length > 3)
            label = label[..3];

        using var bmp = new Bitmap(size, size);
        using (var g = Graphics.FromImage(bmp))
        {
            g.Clear(Color.Transparent);
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using var brush = new SolidBrush(accent);
            FillRoundedRectangle(g, brush, 1, 1, size - 2, size - 2, Math.Max(2, size / 5));
            using var font = new Font("Segoe UI", Math.Max(5f, size * 0.32f), FontStyle.Bold);
            var textSize = g.MeasureString(label, font);
            using var textBrush = new SolidBrush(Color.White);
            g.DrawString(
                label,
                font,
                textBrush,
                (size - textSize.Width) / 2f,
                (size - textSize.Height) / 2f);
        }

        return IconFromBitmap(bmp);
    }

    private static Icon IconFromBitmap(Bitmap bmp)
    {
        var handle = bmp.GetHicon();
        using var temp = Icon.FromHandle(handle);
        var clone = (Icon)temp.Clone();
        DestroyIcon(handle);
        return clone;
    }

    private static void FillRoundedRectangle(Graphics g, Brush brush, int x, int y, int w, int h, int radius)
    {
        using var path = new System.Drawing.Drawing2D.GraphicsPath();
        var d = radius * 2;
        path.AddArc(x, y, d, d, 180, 90);
        path.AddArc(x + w - d, y, d, d, 270, 90);
        path.AddArc(x + w - d, y + h - d, d, d, 0, 90);
        path.AddArc(x, y + h - d, d, d, 90, 90);
        path.CloseFigure();
        g.FillPath(brush, path);
    }
}
