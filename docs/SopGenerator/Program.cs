using System.Runtime.InteropServices;

namespace SopGenerator;

internal static class Program
{
    private const int SlideW = 720;
    private const int SlideH = 405;
    private const int TotalSlides = 15;
    private const string TitleFont = "Microsoft JhengHei";
    private const string BodyFont = "Microsoft JhengHei";

    private static int OfficeBlue => Rgb(43, 87, 154);
    private static int LightBlue => Rgb(210, 230, 240);
    private static int DarkGray => Rgb(51, 51, 51);
    private static int White => Rgb(255, 255, 255);

    private static int Rgb(int r, int g, int b) => r | (g << 8) | (b << 16);

    static int Main(string[] args)
    {
        var outputPath = args.Length > 0
            ? args[0]
            : Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "OfficeLegacyConverter_SOP.pptx");

        outputPath = Path.GetFullPath(outputPath);
        var version = "1.0";
        var date = DateTime.Now.ToString("yyyy/MM/dd");

        dynamic? ppt = null;
        dynamic? pres = null;

        try
        {
            Console.WriteLine("啟動 PowerPoint COM...");
            var pptType = Type.GetTypeFromProgID("PowerPoint.Application")
                ?? throw new InvalidOperationException("找不到 PowerPoint，請確認已安裝。");

            ppt = Activator.CreateInstance(pptType)!;
            ppt.Visible = -1; // 部分環境不允許隱藏 PowerPoint 視窗
            try { ppt.WindowState = 2; } catch { /* ppWindowMinimized — 忽略 */ }

            pres = ppt.Presentations.Add();
            pres.PageSetup.SlideWidth = SlideW;
            pres.PageSetup.SlideHeight = SlideH;

            var n = 0;

            // 1. 封面
            n++;
            var cover = pres.Slides.Add(1, 12);
            SetWhiteBackground(cover, OfficeBlue);
            AddCenteredText(cover, "OfficeLegacyConverter 操作標準程序（SOP）",
                40, 120, 640, 80, 32, true, White);
            AddCenteredText(cover, "Office 舊格式批次轉換工具",
                40, 210, 640, 50, 20, false, LightBlue);
            AddCenteredText(cover, $"版本：{version}\n日期：{date}",
                40, 300, 640, 40, 14, false, White);
            AddSlideNumber(cover, n);

            // 2. 目錄
            n++;
            var toc = NewContentSlide(pres, "目錄");
            AddBodyText(toc,
                "1. 工具目的與適用範圍\n" +
                "2. 系統需求\n" +
                "3. 安裝與啟動\n" +
                "4. 使用者介面總覽\n" +
                "5. 操作 SOP — DOC → DOCX\n" +
                "6. 操作 SOP — XLS → XLSX\n" +
                "7. 操作 SOP — PPT → PPTX\n" +
                "8. 進階選項說明\n" +
                "9. 錯誤排除\n" +
                "10. 技術架構簡介", 75, 18);
            AddSlideNumber(toc, n);

            // 3. 工具目的
            n++;
            var purpose = NewContentSlide(pres, "1. 工具目的與適用範圍");
            AddBodyText(purpose,
                "【工具目的】\n" +
                "• 批次將 Microsoft Office 97–2003 舊格式檔案轉換為 Open XML 格式\n" +
                "• 解決舊版 .doc / .xls / .ppt 在新版 Office 或雲端服務中相容性問題\n" +
                "• 透過 COM 自動化呼叫本機已安裝的 Word、Excel、PowerPoint 進行轉換\n\n" +
                "【適用範圍】\n" +
                "• 需將大量舊格式文件升級為現代格式的使用者\n" +
                "• 檔案來源為 Word 97–2003、Excel 97–2003、PowerPoint 97–2003\n" +
                "• 轉換後檔案儲存於與原始檔相同目錄，副檔名自動變更", 70, 15);
            AddSlideNumber(purpose, n);

            // 4. 格式對照表
            n++;
            var fmt = NewContentSlide(pres, "支援格式對照表");
            var tbl = fmt.Shapes.AddTable(4, 4, 30, 90, 660, 180).Table;
            SetCell(tbl, 1, 1, "來源格式", true, OfficeBlue, White);
            SetCell(tbl, 1, 2, "目標格式", true, OfficeBlue, White);
            SetCell(tbl, 1, 3, "Office 應用程式", true, OfficeBlue, White);
            SetCell(tbl, 1, 4, "COM 格式代碼", true, OfficeBlue, White);
            SetCell(tbl, 2, 1, ".doc");
            SetCell(tbl, 2, 2, ".docx");
            SetCell(tbl, 2, 3, "Microsoft Word");
            SetCell(tbl, 2, 4, "12 (WdFormatXmlDocument)");
            SetCell(tbl, 3, 1, ".xls");
            SetCell(tbl, 3, 2, ".xlsx");
            SetCell(tbl, 3, 3, "Microsoft Excel");
            SetCell(tbl, 3, 4, "51 (XlOpenXmlWorkbook)");
            SetCell(tbl, 4, 1, ".ppt");
            SetCell(tbl, 4, 2, ".pptx");
            SetCell(tbl, 4, 3, "Microsoft PowerPoint");
            SetCell(tbl, 4, 4, "24 (PpSaveAsOpenXmlPresentation)");
            AddBodyText(fmt, "※ 轉換採用 SaveAs / SaveAs2，保留原始檔內容與格式（依 Office 引擎能力）。", 290, 13, 40);
            AddSlideNumber(fmt, n);

            // 5. 系統需求
            n++;
            var req = NewContentSlide(pres, "2. 系統需求");
            AddBodyText(req,
                "【作業系統】\n• Microsoft Windows 10 或以上\n\n" +
                "【執行環境】\n• .NET 10.0（Windows 桌面執行階段）\n• Windows Forms 應用程式（WinExe）\n\n" +
                "【必要軟體】\n• 已安裝 Microsoft Word（處理 .doc → .docx）\n" +
                "• 已安裝 Microsoft Excel（處理 .xls → .xlsx）\n" +
                "• 已安裝 Microsoft PowerPoint（處理 .ppt → .pptx）\n\n" +
                "【COM 自動化說明】\n• 透過 ProgID 啟動 Office 應用程式\n" +
                "• 轉換期間 Office 在背景執行（可選擇顯示視窗以輸入密碼）\n" +
                "• 需確保 Office 已正確安裝且 COM 元件可正常註冊", 70, 14);
            AddSlideNumber(req, n);

            // 6. 安裝與啟動
            n++;
            var install = NewContentSlide(pres, "3. 安裝與啟動");
            AddBodyText(install,
                "【專案路徑】\n• 原始碼目錄：d:\\VS\\Doc2Docx\n\n" +
                "【建置方式】\n1. 開啟終端機，切換至專案目錄\n2. 執行：dotnet build\n" +
                "3. 建置成功後，執行檔位於：bin\\Debug\\net10.0-windows\\OfficeLegacyConverter.exe\n\n" +
                "【執行方式】\n• 方式一：dotnet run（於專案目錄執行）\n" +
                "• 方式二：直接雙擊 OfficeLegacyConverter.exe\n\n" +
                "【注意事項】\n• 首次執行前請確認已安裝 .NET 10 SDK 或執行階段\n" +
                "• 建議以系統管理員權限執行（若 Office COM 權限受限）", 70, 14);
            AddSlideNumber(install, n);

            // 7. UI 總覽
            n++;
            var ui = NewContentSlide(pres, "4. 使用者介面總覽");
            AddBodyText(ui,
                "主視窗標題：「Office 舊格式轉換器」（696 × 572 像素）\n\n" +
                "【區域配置】\n" +
                "① TabControl 分頁列 — 三個分頁：DOC→DOCX、XLS→XLSX、PPT→PPTX\n" +
                "② 檔案清單（ListBox）— 顯示待轉換檔案完整路徑，支援多選與拖放\n" +
                "③ 右側按鈕區 — 加入檔案、加入資料夾、移除選取、全部清除\n" +
                "④ 選項區 — 刪除原始檔、顯示 Office 視窗（兩個核取方塊）\n" +
                "⑤ 操作按鈕 — 開始轉換、取消\n" +
                "⑥ 進度列與狀態標籤 — 顯示轉換進度（N / 總數）\n" +
                "⑦ 日誌文字方塊 — 唯讀，記錄每個檔案的處理結果", 70, 14);
            AddSlideNumber(ui, n);

            // 8. UI Mockup
            n++;
            var mock = NewContentSlide(pres, "4. 使用者介面 — 版面配置圖");
            DrawUiMockup(mock);
            AddBodyText(mock,
                "圖例說明：\n• 切換分頁後，檔案清單與選項文字會依格式自動更新\n" +
                "• 三個 ListBox 重疊於相同位置，僅顯示目前分頁對應的清單\n" +
                "• 轉換進行中會鎖定所有輸入控制項，僅「取消」可按",
                70, 11, 300, 540, 170);
            AddSlideNumber(mock, n);

            // 9-11. SOP slides
            n = AddSopSlide(pres, n, "5. 操作 SOP — DOC → DOCX", "DOC → DOCX", ".doc", "Word");
            n = AddSopSlide(pres, n, "6. 操作 SOP — XLS → XLSX", "XLS → XLSX", ".xls", "Excel");
            n = AddSopSlide(pres, n, "7. 操作 SOP — PPT → PPTX", "PPT → PPTX", ".ppt", "PowerPoint");

            // 12. 進階選項
            n++;
            var adv = NewContentSlide(pres, "8. 進階選項說明");
            AddBodyText(adv,
                "【轉換完成後刪除原始檔】\n• 勾選後，成功轉換的舊格式檔案將被永久刪除\n" +
                "• 點選「開始轉換」時會跳出警告對話框要求再次確認\n• 此操作無法復原，建議先備份重要檔案\n\n" +
                "【顯示 Office 視窗】\n• 預設為背景執行（Word/Excel/PowerPoint 不可見）\n" +
                "• 若檔案設有密碼保護，需勾選此選項以手動輸入密碼\n• 分頁切換時，文字會自動顯示對應應用程式名稱\n\n" +
                "【拖放與批次加入】\n• 支援拖放單一檔案、多個檔案或整個資料夾\n" +
                "• 拖放資料夾時會遞迴搜尋符合副檔名的檔案\n• 重複路徑會自動略過\n\n" +
                "【取消轉換】\n• 轉換進行中可點選「取消」中斷後續檔案處理\n" +
                "• 已完成的檔案不受影響", 70, 13);
            AddSlideNumber(adv, n);

            // 13. 錯誤排除
            n++;
            var err = NewContentSlide(pres, "9. 錯誤排除");
            AddBodyText(err,
                "【找不到 Word / Excel / PowerPoint】\n• 錯誤訊息：「找不到 Microsoft Word，請確認已安裝。」\n" +
                "• 解法：安裝對應 Office 應用程式，或修復 Office 安裝\n\n" +
                "【單一檔案失敗不影響其他檔案】\n• 每個檔案獨立 try/catch 處理，失敗時記錄於日誌\n" +
                "• 日誌格式：「錯誤 - <路徑>：<錯誤訊息>」\n• 批次轉換會繼續處理剩餘檔案\n\n" +
                "【密碼保護檔案】\n• 未勾選「顯示 Office 視窗」時無法輸入密碼，轉換將失敗\n" +
                "• 解法：勾選顯示視窗，於彈出的 Office 對話框輸入密碼\n\n" +
                "【檔案被佔用】\n• 若檔案正被其他程序開啟，Office 無法開啟該檔案\n" +
                "• 解法：關閉佔用該檔案的程式後重試\n\n" +
                "【其他】\n• 磁碟空間不足、路徑過長、權限不足等亦可能導致失敗\n" +
                "• 請查閱日誌區的詳細錯誤訊息以判斷原因", 70, 13);
            AddSlideNumber(err, n);

            // 14. 技術架構
            n++;
            var arch = NewContentSlide(pres, "10. 技術架構簡介");
            AddBodyText(arch,
                "【專案結構】\n• Program.cs — 應用程式進入點，啟動 MainForm\n" +
                "• MainForm.cs / MainForm.Designer.cs — 主視窗 UI 與事件處理\n" +
                "• WordConverter.cs — Word COM 轉換（.doc → .docx）\n" +
                "• ExcelConverter.cs — Excel COM 轉換（.xls → .xlsx）\n" +
                "• PowerPointConverter.cs — PowerPoint COM 轉換（.ppt → .pptx）\n\n" +
                "【轉換流程】\n1. 透過 Type.GetTypeFromProgID 取得 Office COM 類型\n" +
                "2. Activator.CreateInstance 啟動 Office 應用程式\n" +
                "3. 逐一開啟檔案 → SaveAs/SaveAs2 → 關閉文件\n" +
                "4. 可選刪除原始檔，最後 Quit 並 ReleaseComObject\n\n" +
                "【COM 格式代碼】\n• Word：12（WdFormatXmlDocument）\n" +
                "• Excel：51（XlOpenXmlWorkbook）\n• PowerPoint：24（PpSaveAsOpenXmlPresentation）", 70, 13);
            AddSlideNumber(arch, n);

            // 15. 結尾
            n++;
            var end = pres.Slides.Add(pres.Slides.Count + 1, 12);
            SetWhiteBackground(end, OfficeBlue);
            AddCenteredText(end,
                "感謝使用 OfficeLegacyConverter\n\n維護備註：\n" +
                "• 專案路徑：d:\\VS\\Doc2Docx\n• 如有問題請聯繫專案維護人員\n" +
                "• SOP 簡報可透過 docs\\SopGenerator 重新產生",
                40, 130, 640, 140, 18, false, White);
            AddSlideNumber(end, n);

            // 儲存
            var dir = Path.GetDirectoryName(outputPath)!;
            Directory.CreateDirectory(dir);
            if (File.Exists(outputPath)) File.Delete(outputPath);

            pres.SaveAs(outputPath);
            var count = (int)pres.Slides.Count;

            pres.Close();
            ppt.Quit();

            var sizeKb = Math.Round(new FileInfo(outputPath).Length / 1024.0, 1);
            Console.WriteLine($"簡報已產生：{outputPath}");
            Console.WriteLine($"投影片數量：{count}");
            Console.WriteLine($"檔案大小：{sizeKb} KB");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"產生簡報失敗：{ex.Message}");
            return 1;
        }
        finally
        {
            if (pres is not null) Marshal.ReleaseComObject(pres);
            if (ppt is not null)
            {
                try { ppt.Quit(); } catch { /* ignore */ }
                Marshal.ReleaseComObject(ppt);
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private static int AddSopSlide(dynamic pres, int n, string title, string tab, string ext, string app)
    {
        n++;
        var s = NewContentSlide(pres, title);
        AddBodyText(s,
            "1. 啟動 OfficeLegacyConverter.exe\n" +
            $"2. 確認位於「{tab}」分頁\n" +
            "3. 加入待轉換檔案（擇一或多種方式）：\n" +
            $"   • 點選「加入檔案...」選擇一或多個 {ext} 檔案\n" +
            $"   • 點選「加入資料夾...」遞迴掃描資料夾內所有 {ext}\n" +
            "   • 直接拖放檔案或資料夾至檔案清單\n" +
            "4. 檢視清單，必要時使用「移除選取」或「全部清除」\n" +
            $"5. 依需求勾選進階選項（刪除原始檔 / 顯示 {app} 視窗）\n" +
            "6. 點選「開始轉換」\n" +
            "7. 觀察進度列與日誌，等待「全部處理完成。」訊息\n" +
            $"8. 至原始檔目錄確認已產生同名 {ext}x 檔案", 70, 14);
        AddSlideNumber(s, n);
        return n;
    }

    private static void DrawUiMockup(dynamic slide)
    {
        var gray = Rgb(128, 128, 128);
        var lightGray = Rgb(245, 245, 245);
        var medGray = Rgb(224, 224, 224);
        var btnGray = Rgb(232, 232, 232);

        var win = slide.Shapes.AddShape(1, 30, 65, 500, 310);
        win.Fill.ForeColor.RGB = lightGray;
        win.Line.ForeColor.RGB = gray;
        win.Line.Weight = 1.5f;

        var titleBar = slide.Shapes.AddShape(1, 30, 65, 500, 22);
        titleBar.Fill.ForeColor.RGB = OfficeBlue;
        titleBar.Line.Visible = 0;
        SetShapeText(titleBar, "Office 舊格式轉換器", 9, true, White);

        var tabs = new[] { "DOC → DOCX", "XLS → XLSX", "PPT → PPTX" };
        for (var i = 0; i < 3; i++)
        {
            var tab = slide.Shapes.AddShape(1, 32 + i * 110, 87, 105, 18);
            tab.Fill.ForeColor.RGB = i == 0 ? White : medGray;
            tab.Line.ForeColor.RGB = Rgb(170, 170, 170);
            SetShapeText(tab, tabs[i], 7, false, DarkGray, center: true);
        }

        AddSmallLabel(slide, "待轉換檔案（.doc）：", 34, 108);
        var list = slide.Shapes.AddShape(1, 34, 124, 380, 120);
        list.Fill.ForeColor.RGB = White;
        list.Line.ForeColor.RGB = Rgb(170, 170, 170);
        SetShapeText(list, "C:\\Docs\\report.doc\nC:\\Docs\\budget.doc\n（支援拖放檔案或資料夾）", 7, false, Rgb(85, 85, 85), "Consolas");

        var btns = new[] { "加入檔案...", "加入資料夾...", "移除選取", "全部清除" };
        for (var i = 0; i < 4; i++)
        {
            var btn = slide.Shapes.AddShape(1, 420, 124 + i * 28, 100, 22);
            btn.Fill.ForeColor.RGB = btnGray;
            btn.Line.ForeColor.RGB = Rgb(153, 153, 153);
            SetShapeText(btn, btns[i], 7, false, DarkGray, center: true);
        }

        AddSmallLabel(slide, "☐ 轉換完成後刪除原始 .doc 檔案", 34, 252);
        AddSmallLabel(slide, "☐ 顯示 Word 視窗（密碼保護檔案時需要）", 34, 268);

        var convBtn = slide.Shapes.AddShape(1, 34, 290, 80, 24);
        convBtn.Fill.ForeColor.RGB = OfficeBlue;
        convBtn.Line.Visible = 0;
        SetShapeText(convBtn, "開始轉換", 8, true, White, center: true);

        var cancelBtn = slide.Shapes.AddShape(1, 120, 290, 60, 24);
        cancelBtn.Fill.ForeColor.RGB = medGray;
        cancelBtn.Line.ForeColor.RGB = Rgb(153, 153, 153);
        SetShapeText(cancelBtn, "取消", 8, false, DarkGray, center: true);

        var progBg = slide.Shapes.AddShape(1, 34, 322, 486, 12);
        progBg.Fill.ForeColor.RGB = medGray;
        progBg.Line.Visible = 0;
        var progFg = slide.Shapes.AddShape(1, 34, 322, 200, 12);
        progFg.Fill.ForeColor.RGB = OfficeBlue;
        progFg.Line.Visible = 0;

        AddSmallLabel(slide, "2 / 5", 34, 336);

        var logBox = slide.Shapes.AddShape(1, 34, 352, 486, 20);
        logBox.Fill.ForeColor.RGB = White;
        logBox.Line.ForeColor.RGB = Rgb(170, 170, 170);
        SetShapeText(logBox, "開啟 C:\\Docs\\report.doc...\n完成：report.doc → report.docx", 6, false, Rgb(85, 85, 85), "Consolas");
    }

    private static dynamic NewContentSlide(dynamic pres, string title)
    {
        var slide = pres.Slides.Add(pres.Slides.Count + 1, 12);
        SetWhiteBackground(slide, White);
        AddTitleBar(slide, title);
        return slide;
    }

    private static void SetWhiteBackground(dynamic slide, int color)
    {
        slide.FollowMasterBackground = 0;
        slide.Background.Fill.ForeColor.RGB = color;
    }

    private static void AddTitleBar(dynamic slide, string title)
    {
        var bar = slide.Shapes.AddShape(1, 0, 0, SlideW, 56);
        bar.Fill.ForeColor.RGB = OfficeBlue;
        bar.Line.Visible = 0;
        var tr = bar.TextFrame.TextRange;
        tr.Text = title;
        tr.Font.Name = TitleFont;
        tr.Font.Size = 22;
        tr.Font.Bold = -1;
        tr.Font.Color.RGB = White;
        bar.TextFrame.VerticalAnchor = 3;
        bar.TextFrame.MarginLeft = 20;
    }

    private static void AddBodyText(dynamic slide, string text, double top, int size, double height = 300,
        double left = 30, double width = 660)
    {
        var box = slide.Shapes.AddTextbox(1, left, top, width, height);
        box.Line.Visible = 0;
        box.Fill.Visible = 0;
        var tr = box.TextFrame.TextRange;
        tr.Text = text;
        tr.Font.Name = BodyFont;
        tr.Font.Size = size;
        tr.Font.Color.RGB = DarkGray;
        box.TextFrame.WordWrap = -1;
    }

    private static void AddCenteredText(dynamic slide, string text, double left, double top,
        double width, double height, int size, bool bold, int color)
    {
        var box = slide.Shapes.AddTextbox(1, left, top, width, height);
        box.Line.Visible = 0;
        box.Fill.Visible = 0;
        var tr = box.TextFrame.TextRange;
        tr.Text = text;
        tr.Font.Name = TitleFont;
        tr.Font.Size = size;
        tr.Font.Bold = bold ? -1 : 0;
        tr.Font.Color.RGB = color;
        tr.ParagraphFormat.Alignment = 2;
    }

    private static void AddSlideNumber(dynamic slide, int num)
    {
        var box = slide.Shapes.AddTextbox(1, SlideW - 80, SlideH - 28, 60, 20);
        box.Line.Visible = 0;
        box.Fill.Visible = 0;
        var tr = box.TextFrame.TextRange;
        tr.Text = $"{num} / {TotalSlides}";
        tr.Font.Name = BodyFont;
        tr.Font.Size = 10;
        tr.Font.Color.RGB = Rgb(153, 153, 153);
        tr.ParagraphFormat.Alignment = 3;
    }

    private static void SetCell(dynamic table, int row, int col, string text, bool bold = false,
        int? fillColor = null, int? textColor = null)
    {
        var cell = table.Cell(row, col);
        if (fillColor.HasValue)
            cell.Shape.Fill.ForeColor.RGB = fillColor.Value;
        var tr = cell.Shape.TextFrame.TextRange;
        tr.Text = text;
        tr.Font.Name = BodyFont;
        tr.Font.Size = bold ? 13 : 12;
        tr.Font.Bold = bold ? -1 : 0;
        if (textColor.HasValue)
            tr.Font.Color.RGB = textColor.Value;
    }

    private static void SetShapeText(dynamic shape, string text, int size, bool bold, int color,
        string font = "Microsoft JhengHei", bool center = false)
    {
        var tr = shape.TextFrame.TextRange;
        tr.Text = text;
        tr.Font.Name = font;
        tr.Font.Size = size;
        tr.Font.Bold = bold ? -1 : 0;
        tr.Font.Color.RGB = color;
        if (center) tr.ParagraphFormat.Alignment = 2;
    }

    private static void AddSmallLabel(dynamic slide, string text, double left, double top)
    {
        var box = slide.Shapes.AddTextbox(1, left, top, 400, 14);
        box.Line.Visible = 0;
        box.Fill.Visible = 0;
        var tr = box.TextFrame.TextRange;
        tr.Text = text;
        tr.Font.Name = BodyFont;
        tr.Font.Size = 7;
        tr.Font.Color.RGB = DarkGray;
    }
}
