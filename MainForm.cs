namespace OfficeLegacyConverter;

public partial class MainForm : Form
{
    private enum ConversionMode { LegacyUpgrade, OfficePdf, PdfMerge, Markdown }

    private static readonly string MarkdownFileFilter =
        "支援的檔案|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.msg|" +
        "Office / PDF|*.pdf;*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx|" +
        "Outlook 郵件 (*.msg)|*.msg";

    private static readonly string OfficePdfFileFilter =
        "Office 檔案|*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx";

    private CancellationTokenSource? _cts;
    private ConversionProgressForm? _conversionProgressForm;

    private readonly Icon[] _tabIcons = new Icon[4];

    public MainForm()
    {
        InitializeComponent();
        ApplyChrome();
        WireLegacyDragDrop(lstLegacyFiles);
        WireOfficePdfDragDrop(lstDocxFiles);
        WireDragDrop(lstPdfFiles, ".pdf");
        WireMarkdownDragDrop(lstMarkdownFiles);
        UpdateModeUi();
    }

    private void ApplyChrome()
    {
        // Use ApplicationIcon embedded in the exe (not Shell Word icon).
        try
        {
            Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath) ?? Icon;
        }
        catch { /* inherit default from ApplicationIcon */ }

        var assembly = typeof(MainForm).Assembly;
        var version = assembly.GetName().Version?.ToString(3) ?? "unknown";
        var releaseDate = assembly
            .GetCustomAttributes(
                typeof(System.Reflection.AssemblyMetadataAttribute),
                inherit: false)
            .Cast<System.Reflection.AssemblyMetadataAttribute>()
            .FirstOrDefault(a => a.Key == "ReleaseDate")
            ?.Value ?? "unknown";
        lblVersionInfo.Text = $"v{version} · {releaseDate}";
        Text = $"Office 舊格式轉換器 v{version}";

        _tabIcons[0] = ShellIcons.GetOfficeSuiteIcon();
        _tabIcons[1] = ShellIcons.GetOfficeSuiteIcon();
        _tabIcons[2] = ShellIcons.GetSmallIcon(".pdf");
        _tabIcons[3] = ShellIcons.GetSmallIcon(".md");

        foreach (var list in new[] { lstLegacyFiles, lstDocxFiles, lstPdfFiles, lstMarkdownFiles })
            UiTheme.StyleListBox(list);

        UiTheme.StyleSecondaryButton(btnAddFiles);
        UiTheme.StyleSecondaryButton(btnAddFolder);
        UiTheme.StyleSecondaryButton(btnRemove);
        UiTheme.StyleSecondaryButton(btnClear);
        UiTheme.StyleSecondaryButton(btnMoveUp);
        UiTheme.StyleSecondaryButton(btnMoveDown);
        UiTheme.StyleSecondaryButton(btnSelectOutput);
        UiTheme.StyleSecondaryButton(btnCancel);
        UiTheme.StylePrimaryButton(btnConvert, UiTheme.Word);

        // Owner-draw 時隱藏頁面內容區高度，僅作為模式切換列
        tabControl.Appearance = TabAppearance.Normal;
        tabControl.Multiline = false;
    }

    private void TabControl_DrawItem(object? sender, DrawItemEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

        var selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        var accent = UiTheme.AccentForMode(e.Index);
        var bounds = e.Bounds;

        // 填滿標籤背景
        using (var back = new SolidBrush(selected ? Color.White : UiTheme.AppBackground))
            g.FillRectangle(back, bounds);

        if (selected)
        {
            using var fill = new SolidBrush(Color.FromArgb(36, accent));
            g.FillRectangle(fill, bounds);
            using var bar = new SolidBrush(accent);
            g.FillRectangle(bar, bounds.Left, bounds.Bottom - 3, bounds.Width, 3);
        }
        else
        {
            using var edge = new Pen(UiTheme.Border);
            g.DrawLine(edge, bounds.Left, bounds.Bottom - 1, bounds.Right, bounds.Bottom - 1);
        }

        // Windows Shell 檔案類型圖示
        var iconX = bounds.Left + 10;
        var iconY = bounds.Top + (bounds.Height - 16) / 2;
        if (e.Index >= 0 && e.Index < _tabIcons.Length && _tabIcons[e.Index] is { } icon)
            g.DrawIcon(icon, new Rectangle(iconX, iconY, 16, 16));

        var text = tabControl.TabPages[e.Index].Text;
        var textRect = new Rectangle(iconX + 22, bounds.Top, bounds.Width - 34, bounds.Height);
        TextRenderer.DrawText(
            g,
            text,
            selected ? UiTheme.TabFont : UiTheme.UiFont,
            textRect,
            selected ? accent : UiTheme.TextSecondary,
            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
    }

    private void ApplyAccentTheme()
    {
        var accent = UiTheme.AccentForMode(tabControl.SelectedIndex);
        pnlAccent.BackColor = accent;
        lblFiles.ForeColor = accent;
        UiTheme.StylePrimaryButton(btnConvert, accent);
        tabControl.Invalidate();
    }

    private ConversionMode CurrentMode =>
        tabControl.SelectedTab == tabLegacy ? ConversionMode.LegacyUpgrade
        : tabControl.SelectedTab == tabDocxPdf ? ConversionMode.OfficePdf
        : tabControl.SelectedTab == tabPdfMerge ? ConversionMode.PdfMerge
        : ConversionMode.Markdown;

    private ListBox CurrentList =>
        CurrentMode == ConversionMode.LegacyUpgrade ? lstLegacyFiles
        : CurrentMode == ConversionMode.OfficePdf ? lstDocxFiles
        : CurrentMode == ConversionMode.PdfMerge ? lstPdfFiles
        : lstMarkdownFiles;

    private void TabControl_SelectedIndexChanged(object? sender, EventArgs e)
    {
        UpdateModeUi();
        if (CurrentMode == ConversionMode.Markdown)
            LogMarkItDownEnvironmentStatus();
    }

    private void UpdateModeUi()
    {
        var isPdfMerge = CurrentMode == ConversionMode.PdfMerge;
        var isMarkdown = CurrentMode == ConversionMode.Markdown;
        var isLegacy = CurrentMode == ConversionMode.LegacyUpgrade;
        var isOfficePdf = CurrentMode == ConversionMode.OfficePdf;

        lstLegacyFiles.Visible = isLegacy;
        lstDocxFiles.Visible = isOfficePdf;
        lstPdfFiles.Visible = isPdfMerge;
        lstMarkdownFiles.Visible = isMarkdown;

        lblFiles.Text = isPdfMerge
            ? "待合併 PDF 檔案（順序由上而下）"
            : isMarkdown
                ? "待轉換檔案（Office / PDF / MSG）"
                : isOfficePdf
                    ? "待轉換檔案（Office）"
                    : "待轉換檔案（.doc / .xls / .ppt）";

        chkReplace.Visible = !isPdfMerge;
        chkShowOffice.Visible = isLegacy || isOfficePdf || isMarkdown;
        btnMoveUp.Visible = isPdfMerge;
        btnMoveDown.Visible = isPdfMerge;
        lblOutputPath.Visible = isPdfMerge || isMarkdown;
        txtOutputPath.Visible = isPdfMerge || isMarkdown;
        btnSelectOutput.Visible = isPdfMerge || isMarkdown;

        btnConvert.Text = isPdfMerge ? "開始合併" : "開始轉換";

        if (isMarkdown)
        {
            lblOutputPath.Text = "輸出資料夾";
            txtOutputPath.PlaceholderText = "選擇 Markdown 輸出資料夾（含 MSG）";
            btnSelectOutput.Text = "選擇資料夾…";
            chkReplace.Text = "轉換完成後刪除原始檔案";
            chkShowOffice.Text = "顯示 Office 視窗（升級舊格式時）";
        }
        else if (isPdfMerge)
        {
            lblOutputPath.Text = "輸出檔案路徑";
            txtOutputPath.PlaceholderText = "可輸入完整路徑或僅檔名（儲存於第一個 PDF 所在資料夾）";
            btnSelectOutput.Text = "選擇路徑…";
        }
        else if (isOfficePdf)
        {
            chkReplace.Text = "轉換完成後刪除原始檔案";
            chkShowOffice.Text = "顯示 Office 視窗（密碼保護檔案時需要）";
        }
        else if (isLegacy)
        {
            chkReplace.Text = "轉換完成後刪除原始舊格式檔案";
            chkShowOffice.Text = "顯示 Office 視窗（密碼保護檔案時需要）";
        }

        LayoutOptionsPanel(isPdfMerge, isMarkdown);
        ApplyAccentTheme();
    }

    /// <summary>
    /// 依模式重排選項列：Markdown 時輸出路徑 textbox 放在「刪除原始檔」下方第二行，避免與 checkbox 擠在同一列。
    /// </summary>
    private void LayoutOptionsPanel(bool isPdfMerge, bool isMarkdown)
    {
        const int optionsY = 372;
        const int actionsYCompact = 452;
        const int logYCompact = 516;
        const int optionsHeightCompact = 72;
        const int optionsHeightMarkdown = 100;
        const int rowGap = 28;

        if (isMarkdown)
        {
            // 列1：刪除原始檔；列2：輸出路徑 textbox；列3：顯示 Office
            chkReplace.Location = new Point(14, 10);
            lblOutputPath.Location = new Point(14, 10 + rowGap);
            txtOutputPath.Location = new Point(110, 10 + rowGap - 3);
            txtOutputPath.Size = new Size(470, 24);
            btnSelectOutput.Location = new Point(600, 10 + rowGap - 5);
            btnSelectOutput.Size = new Size(114, 28);
            chkShowOffice.Location = new Point(14, 10 + rowGap * 2);

            pnlOptions.Location = new Point(16, optionsY);
            pnlOptions.Size = new Size(728, optionsHeightMarkdown);
            pnlActions.Location = new Point(16, optionsY + optionsHeightMarkdown + 8);
            pnlLog.Location = new Point(16, optionsY + optionsHeightMarkdown + 8 + 64);
            ClientSize = new Size(760, 708);
            MinimumSize = new Size(776, 747);
        }
        else if (isPdfMerge)
        {
            lblOutputPath.Location = new Point(14, 14);
            txtOutputPath.Location = new Point(120, 11);
            txtOutputPath.Size = new Size(470, 24);
            btnSelectOutput.Location = new Point(600, 9);
            btnSelectOutput.Size = new Size(114, 28);

            pnlOptions.Location = new Point(16, optionsY);
            pnlOptions.Size = new Size(728, optionsHeightCompact);
            pnlActions.Location = new Point(16, actionsYCompact);
            pnlLog.Location = new Point(16, logYCompact);
            ClientSize = new Size(760, 680);
            MinimumSize = new Size(776, 719);
        }
        else
        {
            chkReplace.Location = new Point(14, 12);
            chkShowOffice.Location = new Point(14, 40);

            pnlOptions.Location = new Point(16, optionsY);
            pnlOptions.Size = new Size(728, optionsHeightCompact);
            pnlActions.Location = new Point(16, actionsYCompact);
            pnlLog.Location = new Point(16, logYCompact);
            ClientSize = new Size(760, 680);
            MinimumSize = new Size(776, 719);
        }
    }

    private void LogMarkItDownEnvironmentStatus()
    {
        MarkItDownConverter.CheckEnvironment();
        AppendLog(MarkItDownConverter.EnvironmentStatus);
    }

    private void WireDragDrop(ListBox listBox, string extension)
    {
        listBox.DragEnter += (s, e) =>
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        };
        listBox.DragDrop += (s, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] paths)
                return;

            var files = paths
                .SelectMany(path => Directory.Exists(path)
                    ? Directory.EnumerateFiles(path, $"*{extension}", SearchOption.AllDirectories)
                    : [path])
                .Where(path => path.EndsWith(extension, StringComparison.OrdinalIgnoreCase));

            AddFiles(listBox, files, extension);
        };
    }

    private void WireLegacyDragDrop(ListBox listBox)
    {
        listBox.DragEnter += (s, e) =>
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        };
        listBox.DragDrop += (s, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] paths)
                return;

            var files = paths.SelectMany(LegacyOfficeUpgradePipeline.ExpandPaths);
            AddLegacyFiles(listBox, files);
        };
    }

    private void WireOfficePdfDragDrop(ListBox listBox)
    {
        listBox.DragEnter += (s, e) =>
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        };
        listBox.DragDrop += (s, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] paths)
                return;

            var files = paths.SelectMany(ExpandOfficePdfPaths);
            AddOfficePdfFiles(listBox, files);
        };
    }

    private static IEnumerable<string> ExpandOfficePdfPaths(string path)
    {
        if (!Directory.Exists(path))
            return IsOfficePdfSupported(path) ? [path] : [];

        return OfficeToPdfPipeline.SupportedExtensions
            .SelectMany(ext => Directory.EnumerateFiles(path, $"*{ext}", SearchOption.AllDirectories))
            .Where(IsOfficePdfSupported);
    }

    private static bool IsOfficePdfSupported(string path) =>
        OfficeToPdfPipeline.SupportedExtensions.Contains(
            Path.GetExtension(path),
            StringComparer.OrdinalIgnoreCase);

    private void WireMarkdownDragDrop(ListBox listBox)
    {
        listBox.DragEnter += (s, e) =>
        {
            if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
                e.Effect = DragDropEffects.Copy;
        };
        listBox.DragDrop += (s, e) =>
        {
            if (e.Data?.GetData(DataFormats.FileDrop) is not string[] paths)
                return;

            var files = paths.SelectMany(ExpandMarkdownPaths);
            AddMarkdownFiles(listBox, files);
        };
    }

    private static IEnumerable<string> ExpandMarkdownPaths(string path)
    {
        if (!Directory.Exists(path))
            return IsMarkdownSupported(path) ? [path] : [];

        var documents = MarkdownPipeline.SupportedExtensions
            .SelectMany(ext => Directory.EnumerateFiles(path, $"*{ext}", SearchOption.AllDirectories));
        var messages = MsgMarkdownPipeline.ExpandPaths(path);
        return documents.Concat(messages).Where(IsMarkdownSupported);
    }

    private static bool IsMarkdownDocumentSupported(string path) =>
        MarkdownPipeline.SupportedExtensions.Contains(
            Path.GetExtension(path),
            StringComparer.OrdinalIgnoreCase);

    private static bool IsMarkdownSupported(string path) =>
        IsMarkdownDocumentSupported(path) || MsgMarkdownPipeline.IsSupported(path);

    private void BtnAddFiles_Click(object? sender, EventArgs e)
    {
        if (CurrentMode == ConversionMode.Markdown)
        {
            using var dialog = new OpenFileDialog
            {
                Title = "選擇 Office / PDF / MSG 檔案",
                Filter = MarkdownFileFilter,
                Multiselect = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                AddMarkdownFiles(lstMarkdownFiles, dialog.FileNames);
            return;
        }

        if (CurrentMode == ConversionMode.OfficePdf)
        {
            using var dialog = new OpenFileDialog
            {
                Title = "選擇 Office 檔案",
                Filter = OfficePdfFileFilter,
                Multiselect = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                AddOfficePdfFiles(lstDocxFiles, dialog.FileNames);
            return;
        }

        if (CurrentMode == ConversionMode.LegacyUpgrade)
        {
            using var dialog = new OpenFileDialog
            {
                Title = "選擇舊版 Office 檔案",
                Filter = LegacyOfficeUpgradePipeline.FileDialogFilter,
                Multiselect = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                AddLegacyFiles(lstLegacyFiles, dialog.FileNames);
            return;
        }

        using var fileDialog = new OpenFileDialog
        {
            Title = "選擇 PDF 檔案",
            Filter = "PDF (*.pdf)|*.pdf",
            Multiselect = true
        };

        if (fileDialog.ShowDialog() == DialogResult.OK)
            AddFiles(CurrentList, fileDialog.FileNames, ".pdf");
    }

    private void BtnAddFolder_Click(object? sender, EventArgs e)
    {
        if (CurrentMode == ConversionMode.Markdown)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "選擇包含 Office / PDF / MSG 檔案的資料夾",
                UseDescriptionForTitle = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                AddMarkdownFiles(lstMarkdownFiles, ExpandMarkdownPaths(dialog.SelectedPath));
            return;
        }

        if (CurrentMode == ConversionMode.OfficePdf)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "選擇包含 Office 檔案的資料夾",
                UseDescriptionForTitle = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                var files = OfficeToPdfPipeline.SupportedExtensions
                    .SelectMany(ext => Directory.EnumerateFiles(dialog.SelectedPath, $"*{ext}", SearchOption.AllDirectories));
                AddOfficePdfFiles(lstDocxFiles, files);
            }
            return;
        }

        if (CurrentMode == ConversionMode.LegacyUpgrade)
        {
            using var dialog = new FolderBrowserDialog
            {
                Description = "選擇包含 .doc / .xls / .ppt 的資料夾",
                UseDescriptionForTitle = true
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                AddLegacyFiles(lstLegacyFiles, LegacyOfficeUpgradePipeline.ExpandPaths(dialog.SelectedPath));
            return;
        }

        using var folderDialog = new FolderBrowserDialog
        {
            Description = "選擇包含 .pdf 檔案的資料夾",
            UseDescriptionForTitle = true
        };

        if (folderDialog.ShowDialog() == DialogResult.OK)
        {
            var files = Directory.EnumerateFiles(folderDialog.SelectedPath, "*.pdf", SearchOption.AllDirectories);
            AddFiles(CurrentList, files, ".pdf");
        }
    }

    private void BtnRemove_Click(object? sender, EventArgs e)
    {
        var list = CurrentList;
        foreach (var item in list.SelectedItems.Cast<string>().ToList())
            list.Items.Remove(item);
    }

    private void BtnClear_Click(object? sender, EventArgs e) => CurrentList.Items.Clear();

    private void BtnMoveUp_Click(object? sender, EventArgs e)
    {
        var list = CurrentList;
        var index = list.SelectedIndex;
        if (index <= 0)
            return;

        var item = list.Items[index]!;
        list.Items.RemoveAt(index);
        list.Items.Insert(index - 1, item);
        list.SelectedIndex = index - 1;
    }

    private void BtnMoveDown_Click(object? sender, EventArgs e)
    {
        var list = CurrentList;
        var index = list.SelectedIndex;
        if (index < 0 || index >= list.Items.Count - 1)
            return;

        var item = list.Items[index]!;
        list.Items.RemoveAt(index);
        list.Items.Insert(index + 1, item);
        list.SelectedIndex = index + 1;
    }

    private void BtnSelectOutput_Click(object? sender, EventArgs e)
    {
        if (CurrentMode == ConversionMode.Markdown)
        {
            string? initialDir = null;
            if (!string.IsNullOrWhiteSpace(txtOutputPath.Text) && Directory.Exists(txtOutputPath.Text.Trim()))
                initialDir = txtOutputPath.Text.Trim();
            else if (CurrentList.Items.Count > 0)
                initialDir = Path.GetDirectoryName(CurrentList.Items[0]!.ToString()!);

            using var dialog = new FolderBrowserDialog
            {
                Description = "選擇 Markdown 輸出資料夾",
                UseDescriptionForTitle = true,
                SelectedPath = initialDir ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dialog.ShowDialog() == DialogResult.OK)
                txtOutputPath.Text = dialog.SelectedPath;
            return;
        }

        var defaultFileName = "merged.pdf";
        string? pdfInitialDir = null;
        if (lstPdfFiles.Items.Count > 0)
            pdfInitialDir = Path.GetDirectoryName(lstPdfFiles.Items[0]!.ToString()!);

        using var saveDialog = new SaveFileDialog
        {
            Title = "選擇輸出 PDF 檔案",
            Filter = "PDF (*.pdf)|*.pdf",
            DefaultExt = "pdf",
            FileName = defaultFileName,
            InitialDirectory = pdfInitialDir ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
        };

        if (saveDialog.ShowDialog() == DialogResult.OK)
            txtOutputPath.Text = saveDialog.FileName;
    }

    private void AddFiles(ListBox listBox, IEnumerable<string> paths, string extension)
    {
        var existing = listBox.Items.Cast<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);
        var wasEmpty = listBox.Items.Count == 0;
        string? firstNewPath = null;

        foreach (var path in paths)
        {
            if (!path.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
                continue;
            if (existing.Add(path))
            {
                listBox.Items.Add(path);
                firstNewPath ??= path;
            }
        }

        if (extension == ".pdf" && wasEmpty && listBox.Items.Count > 0 && string.IsNullOrWhiteSpace(txtOutputPath.Text))
        {
            var firstPath = firstNewPath ?? listBox.Items[0]!.ToString()!;
            var dir = Path.GetDirectoryName(firstPath);
            if (!string.IsNullOrEmpty(dir))
                txtOutputPath.Text = Path.Combine(dir, "merged.pdf");
        }
    }

    private void AddLegacyFiles(ListBox listBox, IEnumerable<string> paths)
    {
        var existing = listBox.Items.Cast<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var path in paths)
        {
            if (!LegacyOfficeUpgradePipeline.IsSupported(path))
                continue;
            if (existing.Add(path))
                listBox.Items.Add(path);
        }
    }

    private void AddOfficePdfFiles(ListBox listBox, IEnumerable<string> paths)
    {
        var existing = listBox.Items.Cast<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);

        foreach (var path in paths)
        {
            if (!IsOfficePdfSupported(path))
                continue;
            if (existing.Add(path))
                listBox.Items.Add(path);
        }
    }

    private void AddMarkdownFiles(ListBox listBox, IEnumerable<string> paths)
    {
        var existing = listBox.Items.Cast<string>().ToHashSet(StringComparer.OrdinalIgnoreCase);
        var wasEmpty = listBox.Items.Count == 0;
        string? firstNewPath = null;

        foreach (var path in paths)
        {
            if (!IsMarkdownSupported(path))
                continue;
            if (existing.Add(path))
            {
                listBox.Items.Add(path);
                firstNewPath ??= path;
            }
        }

        SuggestOutputFolderIfEmpty(wasEmpty, listBox, firstNewPath);
    }

    private void SuggestOutputFolderIfEmpty(bool wasEmpty, ListBox listBox, string? firstNewPath)
    {
        if (!wasEmpty || listBox.Items.Count == 0 || !string.IsNullOrWhiteSpace(txtOutputPath.Text))
            return;

        var firstPath = firstNewPath ?? listBox.Items[0]!.ToString()!;
        var dir = Path.GetDirectoryName(firstPath);
        if (!string.IsNullOrEmpty(dir))
            txtOutputPath.Text = dir;
    }

    private static bool TryResolvePdfOutputPath(
        string userInput,
        IReadOnlyList<string> inputPdfFiles,
        out string resolvedPath,
        out string errorMessage)
    {
        resolvedPath = "";
        errorMessage = "";

        var trimmed = userInput.Trim();
        if (string.IsNullOrEmpty(trimmed))
        {
            errorMessage = "請選擇或輸入輸出檔案路徑。";
            return false;
        }

        if (trimmed.EndsWith(Path.DirectorySeparatorChar)
            || trimmed.EndsWith(Path.AltDirectorySeparatorChar)
            || Directory.Exists(trimmed))
        {
            errorMessage = "請指定輸出檔案名稱，不能只輸入資料夾路徑。";
            return false;
        }

        string fullPath;
        if (Path.IsPathRooted(trimmed))
        {
            fullPath = Path.GetFullPath(trimmed);
        }
        else
        {
            if (inputPdfFiles.Count == 0)
            {
                errorMessage = "請先加入 PDF 檔案，或指定完整輸出路徑。";
                return false;
            }

            var dir = Path.GetDirectoryName(inputPdfFiles[0]);
            if (string.IsNullOrEmpty(dir))
            {
                errorMessage = "無法判斷輸出資料夾，請指定完整路徑。";
                return false;
            }

            fullPath = Path.GetFullPath(Path.Combine(dir, trimmed));
        }

        if (!fullPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            fullPath += ".pdf";

        var parentDir = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(parentDir))
        {
            errorMessage = "輸出路徑無效。";
            return false;
        }

        if (!Directory.Exists(parentDir))
        {
            errorMessage = $"輸出資料夾不存在：\n{parentDir}";
            return false;
        }

        resolvedPath = fullPath;
        return true;
    }

    private async void BtnConvert_Click(object? sender, EventArgs e)
    {
        if (CurrentMode == ConversionMode.PdfMerge)
        {
            await RunPdfMergeAsync();
            return;
        }

        if (CurrentMode == ConversionMode.Markdown)
        {
            await RunMarkdownConversionAsync();
            return;
        }

        var list = CurrentList;

        if (list.Items.Count == 0)
        {
            var emptyMessage = CurrentMode == ConversionMode.OfficePdf
                ? "請先加入至少一個 Office 檔案。"
                : "請先加入至少一個 .doc / .xls / .ppt 檔案。";
            MessageBox.Show(emptyMessage, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (chkReplace.Checked)
        {
            var replaceMessage = CurrentMode == ConversionMode.OfficePdf
                ? "轉換完成後將刪除原始檔案，此操作無法復原。\n確定要繼續嗎？"
                : "轉換完成後將刪除原始舊格式檔案，此操作無法復原。\n確定要繼續嗎？";
            var result = MessageBox.Show(
                replaceMessage,
                "確認刪除",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
                return;
        }

        var files = list.Items.Cast<string>().ToList();
        var mode = CurrentMode;
        var replace = chkReplace.Checked;
        var showOffice = chkShowOffice.Checked;
        SetConvertingState(true);
        txtLog.Clear();
        progressBar.Maximum = files.Count;
        progressBar.Value = 0;
        ShowConversionProgressForm(files.Count);

        _cts = new CancellationTokenSource();
        var token = _cts.Token;
        var progress = new Progress<(int current, int total, string message)>(UpdateProgress);

        try
        {
            // Excel/Office COM 需要 STA 執行緒；不能用 Task.Run（執行緒池為 MTA，interop 會變慢甚至死結）。
            var result = await RunOnStaThreadAsync(() => mode switch
            {
                ConversionMode.LegacyUpgrade => LegacyOfficeUpgradePipeline.ConvertFiles(files, replace, showOffice, progress, token),
                _ => OfficeToPdfPipeline.ConvertFiles(files, replace, showOffice, progress, token),
            });

            AppendLog("全部處理完成。");
            lblStatus.Text = result.FailureCount > 0 ? "完成（含失敗）" : "完成";
            ShowBatchCompletionSummary(result);
        }
        catch (OperationCanceledException)
        {
            AppendLog("已取消。");
            lblStatus.Text = "已取消";
            MessageBox.Show("已取消轉換。", "已取消", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"發生錯誤：{ex.Message}");
            lblStatus.Text = "錯誤";
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            CloseConversionProgressForm();
            _cts?.Dispose();
            _cts = null;
            SetConvertingState(false);
        }
    }

    /// <summary>
    /// 在專用的 STA 背景執行緒上執行 Office COM 工作，並以 Task 形式回傳結果。
    /// await 會在 UI 執行緒上接續（按鈕還原、MessageBox 都在 UI 執行緒）。
    /// </summary>
    private static Task<T> RunOnStaThreadAsync<T>(Func<T> func)
    {
        var tcs = new TaskCompletionSource<T>();
        var thread = new Thread(() =>
        {
            try
            {
                tcs.SetResult(func());
            }
            catch (OperationCanceledException oce)
            {
                tcs.TrySetCanceled(oce.CancellationToken);
            }
            catch (Exception ex)
            {
                tcs.TrySetException(ex);
            }
        });
        thread.SetApartmentState(ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();
        return tcs.Task;
    }

    private async Task RunMarkdownConversionAsync()
    {
        var files = lstMarkdownFiles.Items.Cast<string>().ToList();
        var outputFolder = txtOutputPath.Text.Trim();

        if (files.Count == 0)
        {
            MessageBox.Show("請先加入至少一個檔案（Office / PDF / MSG）。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (string.IsNullOrWhiteSpace(outputFolder))
        {
            MessageBox.Show("請選擇輸出資料夾。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var resolvedFolder = Path.GetFullPath(outputFolder);
        if (!Directory.Exists(resolvedFolder))
        {
            MessageBox.Show($"輸出資料夾不存在：\n{resolvedFolder}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        txtOutputPath.Text = resolvedFolder;

        var documentFiles = files.Where(IsMarkdownDocumentSupported).ToList();
        var msgFiles = files.Where(MsgMarkdownPipeline.IsSupported).ToList();

        var needsMarkItDown = documentFiles.Any(f =>
        {
            var ext = Path.GetExtension(f);
            return !ext.Equals(".docx", StringComparison.OrdinalIgnoreCase)
                && !ext.Equals(".doc", StringComparison.OrdinalIgnoreCase)
                && !ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)
                && !ext.Equals(".xls", StringComparison.OrdinalIgnoreCase);
        });

        if (needsMarkItDown)
        {
            MarkItDownConverter.CheckEnvironment();
            if (!MarkItDownConverter.IsEnvironmentReady)
            {
                MessageBox.Show(MarkItDownConverter.EnvironmentStatus, "MarkItDown 未就緒", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
        }

        if (chkReplace.Checked)
        {
            var result = MessageBox.Show(
                "轉換完成後將刪除原始檔案，此操作無法復原。\n確定要繼續嗎？",
                "確認刪除",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
                return;
        }

        var replace = chkReplace.Checked;
        var showOffice = chkShowOffice.Checked;
        var total = documentFiles.Count + msgFiles.Count;
        SetConvertingState(true);
        txtLog.Clear();

        if (documentFiles.Count > 0)
        {
            if (needsMarkItDown)
                AppendLog(MarkItDownConverter.EnvironmentStatus);
            else
                AppendLog("Open XML 路徑：Word/Excel 轉換無需 MarkItDown");
        }

        if (msgFiles.Count > 0)
        {
            AppendLog("MSG → Markdown（MsgReader + ReverseMarkdown）");
            AppendLog("附件將寫入 attachments/{mail_id}/");
            AppendLog("Mail analysis：此步驟只做證據擷取，claim ledger、thread 與附件內容核對仍待執行");
            if (Path.GetFileName(resolvedFolder.TrimEnd(
                    Path.DirectorySeparatorChar,
                    Path.AltDirectorySeparatorChar))
                .Equals("extracted", StringComparison.OrdinalIgnoreCase))
            {
                AppendLog("Mail analysis：原始 .msg 將保留至同層 raw/ 資料夾");
            }
        }

        AppendLog($"輸出資料夾：{resolvedFolder}");
        progressBar.Maximum = total;
        progressBar.Value = 0;
        ShowConversionProgressForm(total);

        _cts = new CancellationTokenSource();
        var token = _cts.Token;

        try
        {
            BatchResult? msgResult = null;
            var completed = 0;

            if (documentFiles.Count > 0)
            {
                var docProgress = new Progress<(int current, int total, string message)>(p =>
                    UpdateProgress((completed + p.current, total, p.message)));

                await Task.Run(() =>
                    MarkdownPipeline.ConvertFiles(documentFiles, resolvedFolder, replace, showOffice, docProgress, token));
                completed += documentFiles.Count;
            }

            if (msgFiles.Count > 0)
            {
                var msgProgress = new Progress<(int current, int total, string message)>(p =>
                    UpdateProgress((completed + p.current, total, p.message)));

                msgResult = await Task.Run(() =>
                    MsgMarkdownPipeline.ConvertFiles(msgFiles, resolvedFolder, replace, msgProgress, token));
            }

            AppendLog("全部處理完成。文件已套用 YAML 與品質後處理；MSG 僅完成證據擷取。");
            if (msgResult is { FailureCount: > 0 })
            {
                lblStatus.Text = "完成（含失敗）";
                ShowBatchCompletionSummary(msgResult);
            }
            else
            {
                lblStatus.Text = "完成";
                ShowBatchCompletionMessage(total);
            }
        }
        catch (OperationCanceledException)
        {
            AppendLog("已取消。");
            lblStatus.Text = "已取消";
        }
        catch (Exception ex)
        {
            AppendLog($"發生錯誤：{ex.Message}");
            lblStatus.Text = "錯誤";
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            CloseConversionProgressForm();
            _cts?.Dispose();
            _cts = null;
            SetConvertingState(false);
        }
    }

    private async Task RunPdfMergeAsync()
    {
        var files = lstPdfFiles.Items.Cast<string>().ToList();
        var outputPath = txtOutputPath.Text.Trim();

        if (files.Count < 2)
        {
            MessageBox.Show("請先加入至少兩個 PDF 檔案。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (!TryResolvePdfOutputPath(outputPath, files, out var resolvedPath, out var resolveError))
        {
            MessageBox.Show(resolveError, "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        txtOutputPath.Text = resolvedPath;

        if (File.Exists(resolvedPath))
        {
            var result = MessageBox.Show(
                $"輸出檔案已存在：\n{resolvedPath}\n\n確定要覆寫嗎？",
                "確認覆寫",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result != DialogResult.Yes)
                return;
        }

        SetConvertingState(true);
        txtLog.Clear();
        AppendLog($"輸出路徑：{resolvedPath}");
        progressBar.Maximum = files.Count;
        progressBar.Value = 0;
        ShowConversionProgressForm(files.Count);

        _cts = new CancellationTokenSource();
        var progress = new Progress<(int current, int total, string message)>(UpdateProgress);

        try
        {
            await Task.Run(() =>
                PdfMerger.MergeFiles(files, resolvedPath, progress, _cts.Token));

            AppendLog($"已儲存至：{resolvedPath}");
            AppendLog("全部處理完成。");
            lblStatus.Text = "完成";
            ShowBatchCompletionMessage(files.Count);
        }
        catch (OperationCanceledException)
        {
            AppendLog("已取消。");
            lblStatus.Text = "已取消";
        }
        catch (Exception ex)
        {
            AppendLog($"發生錯誤：{ex.Message}");
            lblStatus.Text = "錯誤";
            MessageBox.Show(ex.Message, "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            CloseConversionProgressForm();
            _cts?.Dispose();
            _cts = null;
            SetConvertingState(false);
        }
    }

    private void BtnCancel_Click(object? sender, EventArgs e) => _cts?.Cancel();

    private void UpdateProgress((int current, int total, string message) report)
    {
        if (InvokeRequired)
        {
            BeginInvoke(UpdateProgress, report);
            return;
        }

        progressBar.Value = Math.Min(report.current, progressBar.Maximum);
        lblStatus.Text = $"{report.current} / {report.total}";
        AppendLog(report.message);
        _conversionProgressForm?.UpdateProgress(report.current, report.total, report.message);
    }

    private void ShowConversionProgressForm(int totalFiles)
    {
        CloseConversionProgressForm();
        _conversionProgressForm = new ConversionProgressForm();
        _conversionProgressForm.UpdateProgress(0, totalFiles, "準備中…");
        _conversionProgressForm.Show(this);
    }

    private void CloseConversionProgressForm()
    {
        if (_conversionProgressForm is null)
            return;

        _conversionProgressForm.CloseSafely();
        _conversionProgressForm.Dispose();
        _conversionProgressForm = null;
    }

    private void AppendLog(string message)
    {
        if (InvokeRequired)
        {
            BeginInvoke(AppendLog, message);
            return;
        }

        txtLog.AppendText(message + Environment.NewLine);
    }

    private void ShowBatchCompletionMessage(int fileCount)
    {
        MessageBox.Show(
            $"已成功處理 {fileCount} 個檔案。",
            "完成",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    private void ShowBatchCompletionSummary(BatchResult result)
    {
        var lines = new List<string>
        {
            "轉換完成。",
            $"成功：{result.SuccessCount} 個",
            $"失敗：{result.FailureCount} 個",
        };

        if (result.FailureCount > 0)
        {
            lines.Add("");
            lines.Add("失敗清單：");
            lines.AddRange(result.Failures.Select(f => $"• {f}"));
        }

        var icon = result.FailureCount > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information;
        MessageBox.Show(
            string.Join(Environment.NewLine, lines),
            "完成",
            MessageBoxButtons.OK,
            icon);
    }

    private void SetConvertingState(bool converting)
    {
        btnConvert.Enabled = !converting;
        btnCancel.Enabled = converting;
        btnAddFiles.Enabled = !converting;
        btnAddFolder.Enabled = !converting;
        btnRemove.Enabled = !converting;
        btnClear.Enabled = !converting;
        btnMoveUp.Enabled = !converting;
        btnMoveDown.Enabled = !converting;
        btnSelectOutput.Enabled = !converting;
        txtOutputPath.Enabled = !converting;
        chkReplace.Enabled = !converting;
        chkShowOffice.Enabled = !converting;
        lstLegacyFiles.Enabled = !converting;
        lstDocxFiles.Enabled = !converting;
        lstPdfFiles.Enabled = !converting;
        lstMarkdownFiles.Enabled = !converting;
        tabControl.Enabled = !converting;
    }

    private void MainForm_Load(object sender, EventArgs e)
    {

    }
}
