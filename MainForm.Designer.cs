namespace OfficeLegacyConverter;

partial class MainForm
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    private void InitializeComponent()
    {
        components = new System.ComponentModel.Container();
        pnlHeader = new Panel();
        lblAppTitle = new Label();
        lblAppSubtitle = new Label();
        lblVersionInfo = new Label();
        pnlAccent = new Panel();
        tabControl = new TabControl();
        tabLegacy = new TabPage();
        tabDocxPdf = new TabPage();
        tabPdfMerge = new TabPage();
        tabMarkdown = new TabPage();
        pnlContent = new Panel();
        lblFiles = new Label();
        lblDropHint = new Label();
        lstLegacyFiles = new ListBox();
        lstDocxFiles = new ListBox();
        lstPdfFiles = new ListBox();
        lstMarkdownFiles = new ListBox();
        pnlSideActions = new Panel();
        btnAddFiles = new Button();
        btnAddFolder = new Button();
        btnRemove = new Button();
        btnClear = new Button();
        btnMoveUp = new Button();
        btnMoveDown = new Button();
        pnlOptions = new Panel();
        chkReplace = new CheckBox();
        chkShowOffice = new CheckBox();
        lblOutputPath = new Label();
        txtOutputPath = new TextBox();
        btnSelectOutput = new Button();
        pnlActions = new Panel();
        btnConvert = new Button();
        btnCancel = new Button();
        progressBar = new ProgressBar();
        lblStatus = new Label();
        pnlLog = new Panel();
        lblLogTitle = new Label();
        txtLog = new TextBox();
        toolTip = new ToolTip(components);

        pnlHeader.SuspendLayout();
        tabControl.SuspendLayout();
        pnlContent.SuspendLayout();
        pnlSideActions.SuspendLayout();
        pnlOptions.SuspendLayout();
        pnlActions.SuspendLayout();
        pnlLog.SuspendLayout();
        SuspendLayout();

        // 
        // pnlHeader
        // 
        pnlHeader.BackColor = UiTheme.HeaderDark;
        pnlHeader.Controls.Add(lblAppTitle);
        pnlHeader.Controls.Add(lblAppSubtitle);
        pnlHeader.Controls.Add(lblVersionInfo);
        pnlHeader.Controls.Add(pnlAccent);
        pnlHeader.Dock = DockStyle.Top;
        pnlHeader.Location = new Point(0, 0);
        pnlHeader.Name = "pnlHeader";
        pnlHeader.Size = new Size(760, 64);
        pnlHeader.TabIndex = 0;
        // 
        // lblAppTitle
        // 
        lblAppTitle.AutoSize = true;
        lblAppTitle.Font = UiTheme.TitleFont;
        lblAppTitle.ForeColor = Color.White;
        lblAppTitle.Location = new Point(20, 12);
        lblAppTitle.Name = "lblAppTitle";
        lblAppTitle.Size = new Size(200, 25);
        lblAppTitle.TabIndex = 0;
        lblAppTitle.Text = "Office 舊格式轉換器";
        // 
        // lblAppSubtitle
        // 
        lblAppSubtitle.AutoSize = true;
        lblAppSubtitle.Font = UiTheme.SubtitleFont;
        lblAppSubtitle.ForeColor = Color.FromArgb(0xB8, 0xC4, 0xD4);
        lblAppSubtitle.Location = new Point(22, 38);
        lblAppSubtitle.Name = "lblAppSubtitle";
        lblAppSubtitle.Size = new Size(340, 15);
        lblAppSubtitle.TabIndex = 1;
        lblAppSubtitle.Text = "批次升級、轉 PDF、合併、Markdown 與郵件匯出";
        // 
        // lblVersionInfo
        // 
        lblVersionInfo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
        lblVersionInfo.Font = UiTheme.SubtitleFont;
        lblVersionInfo.ForeColor = Color.FromArgb(0xB8, 0xC4, 0xD4);
        lblVersionInfo.Location = new Point(500, 20);
        lblVersionInfo.Name = "lblVersionInfo";
        lblVersionInfo.Size = new Size(238, 22);
        lblVersionInfo.TabIndex = 2;
        lblVersionInfo.Text = "v1.1.1 · 2026-09-09";
        lblVersionInfo.TextAlign = ContentAlignment.MiddleRight;
        // 
        // pnlAccent
        // 
        pnlAccent.BackColor = UiTheme.Word;
        pnlAccent.Dock = DockStyle.Bottom;
        pnlAccent.Location = new Point(0, 60);
        pnlAccent.Name = "pnlAccent";
        pnlAccent.Size = new Size(760, 4);
        pnlAccent.TabIndex = 2;
        // 
        // tabControl
        // 
        tabControl.Controls.Add(tabLegacy);
        tabControl.Controls.Add(tabDocxPdf);
        tabControl.Controls.Add(tabPdfMerge);
        tabControl.Controls.Add(tabMarkdown);
        tabControl.DrawMode = TabDrawMode.OwnerDrawFixed;
        tabControl.Font = UiTheme.TabFont;
        tabControl.ItemSize = new Size(170, 40);
        tabControl.Location = new Point(16, 76);
        tabControl.Name = "tabControl";
        tabControl.SelectedIndex = 0;
        tabControl.Size = new Size(728, 44);
        tabControl.SizeMode = TabSizeMode.Fixed;
        tabControl.TabIndex = 1;
        tabControl.DrawItem += TabControl_DrawItem;
        tabControl.SelectedIndexChanged += TabControl_SelectedIndexChanged;
        // 
        // tabLegacy
        // 
        tabLegacy.Location = new Point(4, 44);
        tabLegacy.Name = "tabLegacy";
        tabLegacy.Padding = new Padding(3);
        tabLegacy.Size = new Size(720, 0);
        tabLegacy.TabIndex = 0;
        tabLegacy.Text = "舊格式升級";
        tabLegacy.UseVisualStyleBackColor = true;
        // 
        // tabDocxPdf
        // 
        tabDocxPdf.Location = new Point(4, 44);
        tabDocxPdf.Name = "tabDocxPdf";
        tabDocxPdf.Padding = new Padding(3);
        tabDocxPdf.Size = new Size(720, 0);
        tabDocxPdf.TabIndex = 1;
        tabDocxPdf.Text = "Office → PDF";
        tabDocxPdf.UseVisualStyleBackColor = true;
        // 
        // tabPdfMerge
        // 
        tabPdfMerge.Location = new Point(4, 44);
        tabPdfMerge.Name = "tabPdfMerge";
        tabPdfMerge.Padding = new Padding(3);
        tabPdfMerge.Size = new Size(720, 0);
        tabPdfMerge.TabIndex = 2;
        tabPdfMerge.Text = "PDF 合併";
        tabPdfMerge.UseVisualStyleBackColor = true;
        // 
        // tabMarkdown
        // 
        tabMarkdown.Location = new Point(4, 44);
        tabMarkdown.Name = "tabMarkdown";
        tabMarkdown.Padding = new Padding(3);
        tabMarkdown.Size = new Size(720, 0);
        tabMarkdown.TabIndex = 3;
        tabMarkdown.Text = "轉 Markdown";
        tabMarkdown.UseVisualStyleBackColor = true;
        // 
        // pnlContent
        // 
        pnlContent.BackColor = UiTheme.Surface;
        pnlContent.BorderStyle = BorderStyle.FixedSingle;
        pnlContent.Controls.Add(lblFiles);
        pnlContent.Controls.Add(lblDropHint);
        pnlContent.Controls.Add(lstLegacyFiles);
        pnlContent.Controls.Add(lstDocxFiles);
        pnlContent.Controls.Add(lstPdfFiles);
        pnlContent.Controls.Add(lstMarkdownFiles);
        pnlContent.Controls.Add(pnlSideActions);
        pnlContent.Location = new Point(16, 128);
        pnlContent.Name = "pnlContent";
        pnlContent.Size = new Size(728, 236);
        pnlContent.TabIndex = 2;
        // 
        // lblFiles
        // 
        lblFiles.AutoSize = true;
        lblFiles.Font = new Font("Segoe UI Semibold", 9.5f, FontStyle.Bold);
        lblFiles.ForeColor = UiTheme.TextPrimary;
        lblFiles.Location = new Point(14, 12);
        lblFiles.Name = "lblFiles";
        lblFiles.Size = new Size(220, 17);
        lblFiles.TabIndex = 0;
        lblFiles.Text = "待轉換檔案（.doc / .xls / .ppt）";
        // 
        // lblDropHint
        // 
        lblDropHint.AutoSize = true;
        lblDropHint.Font = UiTheme.SubtitleFont;
        lblDropHint.ForeColor = UiTheme.TextSecondary;
        lblDropHint.Location = new Point(250, 14);
        lblDropHint.Name = "lblDropHint";
        lblDropHint.Size = new Size(200, 15);
        lblDropHint.TabIndex = 1;
        lblDropHint.Text = "可拖放檔案或資料夾至此清單";
        // 
        // lstLegacyFiles
        // 
        lstLegacyFiles.AllowDrop = true;
        lstLegacyFiles.FormattingEnabled = true;
        lstLegacyFiles.HorizontalScrollbar = true;
        lstLegacyFiles.Location = new Point(14, 38);
        lstLegacyFiles.Name = "lstLegacyFiles";
        lstLegacyFiles.SelectionMode = SelectionMode.MultiExtended;
        lstLegacyFiles.Size = new Size(580, 180);
        lstLegacyFiles.TabIndex = 2;
        // 
        // lstDocxFiles
        // 
        lstDocxFiles.AllowDrop = true;
        lstDocxFiles.FormattingEnabled = true;
        lstDocxFiles.HorizontalScrollbar = true;
        lstDocxFiles.Location = new Point(14, 38);
        lstDocxFiles.Name = "lstDocxFiles";
        lstDocxFiles.SelectionMode = SelectionMode.MultiExtended;
        lstDocxFiles.Size = new Size(580, 180);
        lstDocxFiles.TabIndex = 3;
        lstDocxFiles.Visible = false;
        // 
        // lstPdfFiles
        // 
        lstPdfFiles.AllowDrop = true;
        lstPdfFiles.FormattingEnabled = true;
        lstPdfFiles.HorizontalScrollbar = true;
        lstPdfFiles.Location = new Point(14, 38);
        lstPdfFiles.Name = "lstPdfFiles";
        lstPdfFiles.SelectionMode = SelectionMode.MultiExtended;
        lstPdfFiles.Size = new Size(580, 180);
        lstPdfFiles.TabIndex = 4;
        lstPdfFiles.Visible = false;
        // 
        // lstMarkdownFiles
        // 
        lstMarkdownFiles.AllowDrop = true;
        lstMarkdownFiles.FormattingEnabled = true;
        lstMarkdownFiles.HorizontalScrollbar = true;
        lstMarkdownFiles.Location = new Point(14, 38);
        lstMarkdownFiles.Name = "lstMarkdownFiles";
        lstMarkdownFiles.SelectionMode = SelectionMode.MultiExtended;
        lstMarkdownFiles.Size = new Size(580, 180);
        lstMarkdownFiles.TabIndex = 5;
        lstMarkdownFiles.Visible = false;
        // 
        // pnlSideActions
        // 
        pnlSideActions.Controls.Add(btnAddFiles);
        pnlSideActions.Controls.Add(btnAddFolder);
        pnlSideActions.Controls.Add(btnRemove);
        pnlSideActions.Controls.Add(btnClear);
        pnlSideActions.Controls.Add(btnMoveUp);
        pnlSideActions.Controls.Add(btnMoveDown);
        pnlSideActions.Location = new Point(604, 38);
        pnlSideActions.Name = "pnlSideActions";
        pnlSideActions.Size = new Size(110, 180);
        pnlSideActions.TabIndex = 6;
        // 
        // btnAddFiles
        // 
        btnAddFiles.Location = new Point(0, 0);
        btnAddFiles.Name = "btnAddFiles";
        btnAddFiles.Size = new Size(110, 30);
        btnAddFiles.TabIndex = 0;
        btnAddFiles.Text = "加入檔案…";
        btnAddFiles.Click += BtnAddFiles_Click;
        // 
        // btnAddFolder
        // 
        btnAddFolder.Location = new Point(0, 36);
        btnAddFolder.Name = "btnAddFolder";
        btnAddFolder.Size = new Size(110, 30);
        btnAddFolder.TabIndex = 1;
        btnAddFolder.Text = "加入資料夾…";
        btnAddFolder.Click += BtnAddFolder_Click;
        // 
        // btnRemove
        // 
        btnRemove.Location = new Point(0, 72);
        btnRemove.Name = "btnRemove";
        btnRemove.Size = new Size(110, 30);
        btnRemove.TabIndex = 2;
        btnRemove.Text = "移除選取";
        btnRemove.Click += BtnRemove_Click;
        // 
        // btnClear
        // 
        btnClear.Location = new Point(0, 108);
        btnClear.Name = "btnClear";
        btnClear.Size = new Size(110, 30);
        btnClear.TabIndex = 3;
        btnClear.Text = "全部清除";
        btnClear.Click += BtnClear_Click;
        // 
        // btnMoveUp
        // 
        btnMoveUp.Location = new Point(0, 144);
        btnMoveUp.Name = "btnMoveUp";
        btnMoveUp.Size = new Size(52, 30);
        btnMoveUp.TabIndex = 4;
        btnMoveUp.Text = "↑";
        btnMoveUp.Visible = false;
        btnMoveUp.Click += BtnMoveUp_Click;
        // 
        // btnMoveDown
        // 
        btnMoveDown.Location = new Point(58, 144);
        btnMoveDown.Name = "btnMoveDown";
        btnMoveDown.Size = new Size(52, 30);
        btnMoveDown.TabIndex = 5;
        btnMoveDown.Text = "↓";
        btnMoveDown.Visible = false;
        btnMoveDown.Click += BtnMoveDown_Click;
        // 
        // pnlOptions
        // 
        pnlOptions.BackColor = UiTheme.Surface;
        pnlOptions.BorderStyle = BorderStyle.FixedSingle;
        pnlOptions.Controls.Add(chkReplace);
        pnlOptions.Controls.Add(chkShowOffice);
        pnlOptions.Controls.Add(lblOutputPath);
        pnlOptions.Controls.Add(txtOutputPath);
        pnlOptions.Controls.Add(btnSelectOutput);
        pnlOptions.Location = new Point(16, 372);
        pnlOptions.Name = "pnlOptions";
        pnlOptions.Size = new Size(728, 72);
        pnlOptions.TabIndex = 3;
        // 
        // chkReplace
        // 
        chkReplace.AutoSize = true;
        chkReplace.Font = UiTheme.UiFont;
        chkReplace.ForeColor = UiTheme.TextPrimary;
        chkReplace.Location = new Point(14, 12);
        chkReplace.Name = "chkReplace";
        chkReplace.Size = new Size(260, 21);
        chkReplace.TabIndex = 0;
        chkReplace.Text = "轉換完成後刪除原始舊格式檔案";
        chkReplace.UseVisualStyleBackColor = true;
        // 
        // chkShowOffice
        // 
        chkShowOffice.AutoSize = true;
        chkShowOffice.Font = UiTheme.UiFont;
        chkShowOffice.ForeColor = UiTheme.TextPrimary;
        chkShowOffice.Location = new Point(14, 40);
        chkShowOffice.Name = "chkShowOffice";
        chkShowOffice.Size = new Size(280, 21);
        chkShowOffice.TabIndex = 1;
        chkShowOffice.Text = "顯示 Office 視窗（密碼保護檔案時需要）";
        chkShowOffice.UseVisualStyleBackColor = true;
        // 
        // lblOutputPath
        // 
        lblOutputPath.AutoSize = true;
        lblOutputPath.Font = UiTheme.UiFont;
        lblOutputPath.ForeColor = UiTheme.TextPrimary;
        lblOutputPath.Location = new Point(14, 14);
        lblOutputPath.Name = "lblOutputPath";
        lblOutputPath.Size = new Size(91, 17);
        lblOutputPath.TabIndex = 2;
        lblOutputPath.Text = "輸出檔案路徑";
        lblOutputPath.Visible = false;
        // 
        // txtOutputPath
        // 
        txtOutputPath.Font = UiTheme.UiFont;
        txtOutputPath.Location = new Point(120, 11);
        txtOutputPath.Name = "txtOutputPath";
        txtOutputPath.PlaceholderText = "可輸入完整路徑或僅檔名（儲存於第一個 PDF 所在資料夾）";
        txtOutputPath.Size = new Size(470, 24);
        txtOutputPath.TabIndex = 3;
        toolTip.SetToolTip(txtOutputPath, "可輸入完整路徑或僅檔名（將儲存於第一個 PDF 所在資料夾）。建議使用「選擇輸出路徑…」按鈕。");
        txtOutputPath.Visible = false;
        // 
        // btnSelectOutput
        // 
        btnSelectOutput.Location = new Point(600, 9);
        btnSelectOutput.Name = "btnSelectOutput";
        btnSelectOutput.Size = new Size(114, 28);
        btnSelectOutput.TabIndex = 4;
        btnSelectOutput.Text = "選擇路徑…";
        toolTip.SetToolTip(btnSelectOutput, "選擇合併後 PDF 的儲存位置");
        btnSelectOutput.Visible = false;
        btnSelectOutput.Click += BtnSelectOutput_Click;
        // 
        // toolTip
        // 
        toolTip.AutoPopDelay = 8000;
        toolTip.InitialDelay = 500;
        toolTip.ReshowDelay = 200;
        // 
        // pnlActions
        // 
        pnlActions.Controls.Add(btnConvert);
        pnlActions.Controls.Add(btnCancel);
        pnlActions.Controls.Add(progressBar);
        pnlActions.Controls.Add(lblStatus);
        pnlActions.Location = new Point(16, 452);
        pnlActions.Name = "pnlActions";
        pnlActions.Size = new Size(728, 56);
        pnlActions.TabIndex = 4;
        // 
        // btnConvert
        // 
        btnConvert.Location = new Point(0, 0);
        btnConvert.Name = "btnConvert";
        btnConvert.Size = new Size(128, 36);
        btnConvert.TabIndex = 0;
        btnConvert.Text = "開始轉換";
        btnConvert.Click += BtnConvert_Click;
        // 
        // btnCancel
        // 
        btnCancel.Enabled = false;
        btnCancel.Location = new Point(136, 0);
        btnCancel.Name = "btnCancel";
        btnCancel.Size = new Size(100, 36);
        btnCancel.TabIndex = 1;
        btnCancel.Text = "取消";
        btnCancel.Click += BtnCancel_Click;
        // 
        // progressBar
        // 
        progressBar.Location = new Point(250, 6);
        progressBar.Name = "progressBar";
        progressBar.Size = new Size(478, 24);
        progressBar.Style = ProgressBarStyle.Continuous;
        progressBar.TabIndex = 2;
        // 
        // lblStatus
        // 
        lblStatus.AutoSize = true;
        lblStatus.Font = UiTheme.SubtitleFont;
        lblStatus.ForeColor = UiTheme.TextSecondary;
        lblStatus.Location = new Point(250, 34);
        lblStatus.Name = "lblStatus";
        lblStatus.Size = new Size(31, 15);
        lblStatus.TabIndex = 3;
        lblStatus.Text = "就緒";
        // 
        // pnlLog
        // 
        pnlLog.BackColor = UiTheme.Surface;
        pnlLog.BorderStyle = BorderStyle.FixedSingle;
        pnlLog.Controls.Add(lblLogTitle);
        pnlLog.Controls.Add(txtLog);
        pnlLog.Location = new Point(16, 516);
        pnlLog.Name = "pnlLog";
        pnlLog.Size = new Size(728, 148);
        pnlLog.TabIndex = 5;
        // 
        // lblLogTitle
        // 
        lblLogTitle.AutoSize = true;
        lblLogTitle.Font = new Font("Segoe UI Semibold", 9f, FontStyle.Bold);
        lblLogTitle.ForeColor = UiTheme.TextSecondary;
        lblLogTitle.Location = new Point(12, 8);
        lblLogTitle.Name = "lblLogTitle";
        lblLogTitle.Size = new Size(55, 15);
        lblLogTitle.TabIndex = 0;
        lblLogTitle.Text = "執行紀錄";
        // 
        // txtLog
        // 
        txtLog.BackColor = Color.FromArgb(0xFA, 0xFB, 0xFC);
        txtLog.BorderStyle = BorderStyle.None;
        txtLog.Font = new Font("Consolas", 9f);
        txtLog.ForeColor = UiTheme.TextPrimary;
        txtLog.Location = new Point(12, 28);
        txtLog.Multiline = true;
        txtLog.Name = "txtLog";
        txtLog.ReadOnly = true;
        txtLog.ScrollBars = ScrollBars.Vertical;
        txtLog.Size = new Size(702, 108);
        txtLog.TabIndex = 1;
        txtLog.WordWrap = false;
        // 
        // MainForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        BackColor = UiTheme.AppBackground;
        ClientSize = new Size(760, 680);
        Controls.Add(pnlLog);
        Controls.Add(pnlActions);
        Controls.Add(pnlOptions);
        Controls.Add(pnlContent);
        Controls.Add(tabControl);
        Controls.Add(pnlHeader);
        Font = UiTheme.UiFont;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        MinimumSize = new Size(776, 719);
        Name = "MainForm";
        StartPosition = FormStartPosition.CenterScreen;
        Text = "Office 舊格式轉換器";
        Load += MainForm_Load;
        pnlHeader.ResumeLayout(false);
        pnlHeader.PerformLayout();
        tabControl.ResumeLayout(false);
        pnlContent.ResumeLayout(false);
        pnlContent.PerformLayout();
        pnlSideActions.ResumeLayout(false);
        pnlOptions.ResumeLayout(false);
        pnlOptions.PerformLayout();
        pnlActions.ResumeLayout(false);
        pnlActions.PerformLayout();
        pnlLog.ResumeLayout(false);
        pnlLog.PerformLayout();
        ResumeLayout(false);
    }

    #endregion

    private Panel pnlHeader;
    private Label lblAppTitle;
    private Label lblAppSubtitle;
    private Label lblVersionInfo;
    private Panel pnlAccent;
    private TabControl tabControl;
    private TabPage tabLegacy;
    private TabPage tabDocxPdf;
    private TabPage tabPdfMerge;
    private TabPage tabMarkdown;
    private Panel pnlContent;
    private Label lblFiles;
    private Label lblDropHint;
    private ListBox lstLegacyFiles;
    private ListBox lstDocxFiles;
    private ListBox lstPdfFiles;
    private ListBox lstMarkdownFiles;
    private Panel pnlSideActions;
    private Button btnAddFiles;
    private Button btnAddFolder;
    private Button btnRemove;
    private Button btnClear;
    private Button btnMoveUp;
    private Button btnMoveDown;
    private Panel pnlOptions;
    private CheckBox chkReplace;
    private CheckBox chkShowOffice;
    private Label lblOutputPath;
    private TextBox txtOutputPath;
    private Button btnSelectOutput;
    private Panel pnlActions;
    private Button btnConvert;
    private Button btnCancel;
    private TextBox txtLog;
    private ProgressBar progressBar;
    private Label lblStatus;
    private Panel pnlLog;
    private Label lblLogTitle;
    private ToolTip toolTip;
}
