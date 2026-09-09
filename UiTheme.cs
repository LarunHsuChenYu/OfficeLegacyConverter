namespace OfficeLegacyConverter;

/// <summary>Microsoft Office 品牌色與共用 UI 樣式。</summary>
internal static class UiTheme
{
    public static readonly Color Word = Color.FromArgb(0x2B, 0x57, 0x9A);
    public static readonly Color Excel = Color.FromArgb(0x21, 0x73, 0x46);
    public static readonly Color PowerPoint = Color.FromArgb(0xC4, 0x3E, 0x1C);
    public static readonly Color Pdf = Color.FromArgb(0xE5, 0x25, 0x2A);
    public static readonly Color OfficePdf = Color.FromArgb(0x5B, 0x2C, 0x6F);
    public static readonly Color Markdown = Color.FromArgb(0x0E, 0x6E, 0x8A);
    public static readonly Color Outlook = Color.FromArgb(0x00, 0x73, 0xC6);

    public static readonly Color AppBackground = Color.FromArgb(0xF3, 0xF5, 0xF8);
    public static readonly Color Surface = Color.White;
    public static readonly Color Border = Color.FromArgb(0xD0, 0xD7, 0xDE);
    public static readonly Color TextPrimary = Color.FromArgb(0x1F, 0x23, 0x28);
    public static readonly Color TextSecondary = Color.FromArgb(0x5C, 0x65, 0x70);
    public static readonly Color HeaderDark = Color.FromArgb(0x1B, 0x2A, 0x41);

    public static Font TitleFont { get; } = new("Segoe UI Semibold", 14f, FontStyle.Bold);
    public static Font SubtitleFont { get; } = new("Segoe UI", 9f);
    public static Font UiFont { get; } = new("Segoe UI", 9.5f);
    public static Font TabFont { get; } = new("Segoe UI Semibold", 8.5f, FontStyle.Bold);
    public static Font ButtonFont { get; } = new("Segoe UI Semibold", 9.5f);

    public static Color AccentForMode(int tabIndex) => tabIndex switch
    {
        0 => Word,       // 舊格式升級（合併 DOC/XLS/PPT）
        1 => OfficePdf,  // Office → PDF
        2 => Pdf,        // PDF 合併
        3 => Markdown,   // 轉 Markdown（含 MSG）
        _ => Word
    };

    public static Color Lerp(Color a, Color b, float t)
    {
        t = Math.Clamp(t, 0f, 1f);
        return Color.FromArgb(
            (int)(a.A + (b.A - a.A) * t),
            (int)(a.R + (b.R - a.R) * t),
            (int)(a.G + (b.G - a.G) * t),
            (int)(a.B + (b.B - a.B) * t));
    }

    public static void StyleSecondaryButton(Button button)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderColor = Border;
        button.FlatAppearance.BorderSize = 1;
        button.FlatAppearance.MouseOverBackColor = Color.FromArgb(0xEE, 0xF2, 0xF6);
        button.FlatAppearance.MouseDownBackColor = Color.FromArgb(0xE2, 0xE8, 0xEF);
        button.BackColor = Surface;
        button.ForeColor = TextPrimary;
        button.Font = ButtonFont;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
    }

    public static void StylePrimaryButton(Button button, Color accent)
    {
        button.FlatStyle = FlatStyle.Flat;
        button.FlatAppearance.BorderSize = 0;
        button.FlatAppearance.MouseOverBackColor = Lerp(accent, Color.White, 0.12f);
        button.FlatAppearance.MouseDownBackColor = Lerp(accent, Color.Black, 0.12f);
        button.BackColor = accent;
        button.ForeColor = Color.White;
        button.Font = ButtonFont;
        button.Cursor = Cursors.Hand;
        button.UseVisualStyleBackColor = false;
    }

    public static void StyleListBox(ListBox listBox)
    {
        listBox.BorderStyle = BorderStyle.FixedSingle;
        listBox.BackColor = Surface;
        listBox.ForeColor = TextPrimary;
        listBox.Font = UiFont;
        listBox.IntegralHeight = false;
    }
}
