using System.Runtime.InteropServices;
using System.Text;

namespace OfficeLegacyConverter;

internal static class ExcelComHelper
{
    private const int XlTypePdf = 0;
    private const int XlCalculationManual = -4135;
    private const int XlSheetVisible = -1;
    private const int XlSheetHidden = 0;

    /// <summary>
    /// 建立獨立 Excel 執行個體供 PDF 匯出，避免附加至使用者已開啟且忙碌的 Excel（RPC_E_SERVERCALL_RETRYLATER）。
    /// </summary>
    public static dynamic CreateDedicatedExcelForPdfExport()
    {
        var excelType = Type.GetTypeFromProgID("Excel.Application")
            ?? throw new InvalidOperationException("找不到 Microsoft Excel，請確認已安裝。");

        return Activator.CreateInstance(excelType)!;
    }

    /// <summary>
    /// 為 PDF 匯出開啟工作簿：不重用使用者已開啟的檔案，優先可寫入開啟，失敗則複製至安全暫存路徑。
    /// ExportAsFixedFormat 要求工作簿可儲存；唯讀或未儲存狀態會觸發 COM 1004。
    /// </summary>
    /// <returns>工作簿、是否由本程式開啟（可安全 Close）、暫存複本路徑（需清理）。</returns>
    public static (dynamic Workbook, bool OpenedByUs, string? TempCopyPath) OpenWorkbookForPdfExport(
        object excel,
        string inputPath)
    {
        dynamic xl = excel;

        try
        {
            // UpdateLinks=0, ReadOnly=false, IgnoreReadOnlyRecommended=true
            var workbook = xl.Workbooks.Open(
                inputPath,
                0,
                false,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                true);
            return (workbook, true, null);
        }
        catch (COMException ex) when (IsFileAccessComError(ex))
        {
            var tempCopy = CopyToTempSafe(inputPath);
            try
            {
                var workbook = xl.Workbooks.Open(
                    tempCopy,
                    0,
                    false,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true);
                return (workbook, true, tempCopy);
            }
            catch (Exception openEx)
            {
                TryDeleteFile(tempCopy);
                throw WrapExportException("開啟暫存複本", openEx);
            }
        }
    }

    /// <summary>建立純 ASCII 暫存 PDF 路徑（Excel COM 對含 &amp; 等特殊字元路徑較敏感）。</summary>
    public static string CreateSafeTempPdfPath()
    {
        var path = Path.GetFullPath(Path.Combine(
            Path.GetTempPath(),
            "OfficeLegacyConverter",
            Guid.NewGuid().ToString("N") + ".pdf"));
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        return path;
    }

    public static void ConfigureExcelForExport(dynamic excel)
    {
        excel.DisplayAlerts = false;
        excel.ScreenUpdating = false;
        excel.EnableEvents = false;

        try
        {
            excel.Calculation = XlCalculationManual;
        }
        catch (COMException)
        {
            // 部分 Excel 版本或狀態下可能無法切換
        }
    }

    /// <summary>
    /// 在執行任何匯出策略之前，對工作簿中所有工作表套用列印版面（FitToPagesWide=1、窄邊界、重複標頭列、限縮列印範圍等），
    /// 並把空白工作表暫時隱藏以減少空白頁。SaveCopyAs 產生的副本會繼承來源工作簿目前的 PageSetup 與可見性，
    /// 因此先在來源工作簿套用即可。
    /// </summary>
    /// <returns>被暫時隱藏的工作表清單與其原始 Visible 值，呼叫端應於匯出結束後（workbook.Close 前）以 <see cref="RestoreSheetVisibility"/> 還原。</returns>
    public static List<(dynamic Sheet, int OriginalVisible)> ApplyPrintLayout(
        dynamic excel,
        dynamic workbook,
        int sheetCount,
        IProgress<(int current, int total, string message)>? progress,
        int current,
        int total,
        CancellationToken cancellationToken)
    {
        var hiddenSheets = new List<(dynamic Sheet, int OriginalVisible)>();
        var blankCandidates = new List<(dynamic Sheet, int OriginalVisible)>();
        var hasVisibleNonBlank = false;

        for (var i = 1; i <= sheetCount; i++)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if (i == 1 || i == sheetCount || i % 5 == 0)
                progress?.Report((current, total, $"設定版面中（{i}/{sheetCount}）…"));

            dynamic ws = workbook.Worksheets[i];

            // 圖表或唯讀工作表可能不支援部分屬性，整段以 try/catch 包住以略過。
            try
            {
                dynamic pageSetup = ws.PageSetup;

                pageSetup.Zoom = false;              // 必須先關閉 Zoom，FitToPages 才會生效
                pageSetup.FitToPagesWide = 1;        // 欄位縮到一頁寬
                pageSetup.FitToPagesTall = false;    // 高度不限
                pageSetup.PrintTitleRows = "$1:$1";  // 每頁重複第 1 列當標頭

                // 窄邊界（單位為點，1 inch = 72 points）
                pageSetup.LeftMargin = 18;
                pageSetup.RightMargin = 18;
                pageSetup.TopMargin = 54;
                pageSetup.BottomMargin = 54;
                pageSetup.HeaderMargin = 21.6;
                pageSetup.FooterMargin = 21.6;

                pageSetup.CenterHorizontally = true;

                // 減少空白頁：清除手動分頁，並把列印範圍限縮到實際資料
                try { ws.ResetAllPageBreaks(); } catch { /* 部分工作表不支援 */ }
                try
                {
                    dynamic ur = ws.UsedRange;
                    if (ur is not null)
                        pageSetup.PrintArea = (string)ur.Address;
                }
                catch { /* 部分工作表不支援 */ }
            }
            catch { /* 圖表或唯讀工作表略過版面設定 */ }

            // 空白工作表偵測（僅針對目前可見的工作表），暫存為候選，待確認仍有可見內容後再隱藏。
            try
            {
                if ((int)ws.Visible == XlSheetVisible)
                {
                    if ((int)excel.WorksheetFunction.CountA(ws.UsedRange) == 0)
                        blankCandidates.Add((ws, XlSheetVisible));
                    else
                        hasVisibleNonBlank = true;
                }
            }
            catch { /* 無法判定時保守視為有內容，不隱藏 */ }
        }

        // 至少保留一個可見工作表：唯有仍存在可見且非空白的工作表時，才隱藏空白表。
        if (hasVisibleNonBlank)
        {
            foreach (var (sheet, original) in blankCandidates)
            {
                try
                {
                    sheet.Visible = XlSheetHidden;
                    hiddenSheets.Add((sheet, original));
                }
                catch { /* 隱藏失敗則略過 */ }
            }
        }

        return hiddenSheets;
    }

    /// <summary>
    /// 還原 <see cref="ApplyPrintLayout"/> 暫時隱藏的工作表可見性（須在 workbook.Close 之前呼叫）。
    /// </summary>
    public static void RestoreSheetVisibility(IEnumerable<(dynamic Sheet, int OriginalVisible)> sheets)
    {
        foreach (var (sheet, original) in sheets)
        {
            try { sheet.Visible = original; } catch { /* ignore */ }
        }
    }

    /// <summary>
    /// 嘗試讓工作簿處於可匯出狀態：非唯讀且未儲存時 Save；唯讀或 Save 失敗則 SaveCopyAs。
    /// </summary>
    /// <returns>用於匯出的工作簿、是否應由呼叫端 Close、暫存 xlsx 路徑（需清理）。</returns>
    public static (dynamic Workbook, bool CloseWhenDone, string? TempXlsxPath) PrepareWorkbookForExport(
        dynamic excel,
        dynamic sourceWorkbook,
        string sourcePath,
        bool forceSaveCopy = false)
    {
        if (forceSaveCopy)
            return OpenSaveCopy(excel, sourceWorkbook, sourcePath);

        try
        {
            var readOnly = (bool)sourceWorkbook.ReadOnly;
            if (readOnly)
                return OpenSaveCopy(excel, sourceWorkbook, sourcePath);

            var saved = (bool)sourceWorkbook.Saved;
            if (!saved)
            {
                sourceWorkbook.Save();
                return (sourceWorkbook, false, null);
            }

            return (sourceWorkbook, false, null);
        }
        catch (COMException)
        {
            return OpenSaveCopy(excel, sourceWorkbook, sourcePath);
        }
    }

    /// <summary>
    /// 以 Workbook.ExportAsFixedFormat 匯出整本工作簿（不需 Select 工作表）。
    /// </summary>
    public static void ExportWorkbookAsPdf(dynamic workbook, string pdfPath)
    {
        var normalizedPath = Path.GetFullPath(pdfPath);

        // 參數順序：Type, FileName, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish
        workbook.ExportAsFixedFormat(
            XlTypePdf,
            normalizedPath,
            Type.Missing,
            true,
            Type.Missing,
            Type.Missing,
            Type.Missing,
            false,
            Type.Missing);
    }

    /// <summary>
    /// 逐工作表匯出 PDF 後合併（適用於超大工作簿或整本匯出失敗時）。
    /// </summary>
    public static void ExportSheetsAsPdfAndMerge(
        dynamic workbook,
        int sheetCount,
        string outputPdfPath,
        IProgress<(int current, int total, string message)>? progress,
        int progressCurrent,
        int progressTotal,
        CancellationToken cancellationToken)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "OfficeLegacyConverter", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var sheetPdfs = new List<string>(sheetCount);

        try
        {
            for (var i = 1; i <= sheetCount; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();

                progress?.Report((progressCurrent, progressTotal, $"匯出工作表 {i}/{sheetCount}..."));

                var sheetPdf = Path.Combine(tempDir, $"sheet_{i:000}.pdf");
                dynamic? sheet = null;

                try
                {
                    sheet = workbook.Worksheets[i];

                    // 跳過隱藏工作表（含 ApplyPrintLayout 暫時隱藏的空白表）：
                    // 對隱藏表呼叫 ExportAsFixedFormat 會失敗（-1 = xlSheetVisible）。
                    if ((int)sheet.Visible != XlSheetVisible)
                        continue;

                    ExportWorksheetAsPdf(sheet, sheetPdf);

                    if (!File.Exists(sheetPdf) || new FileInfo(sheetPdf).Length == 0)
                        throw new InvalidOperationException($"工作表 {i} 未產生有效 PDF：{sheetPdf}");

                    sheetPdfs.Add(sheetPdf);
                }
                finally
                {
                    if (sheet is not null)
                        Marshal.ReleaseComObject(sheet);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report((progressCurrent, progressTotal, $"合併 {sheetPdfs.Count} 個工作表 PDF..."));
            PdfMerger.MergeFilesOrThrow(sheetPdfs, outputPdfPath);
        }
        finally
        {
            TryDeleteDirectory(tempDir);
        }
    }

    public static void PublishTempPdf(string tempPdfPath, string pdfPath)
    {
        if (!File.Exists(tempPdfPath))
            throw new FileNotFoundException($"暫存 PDF 不存在：{tempPdfPath}");

        var pdfDir = Path.GetDirectoryName(pdfPath);
        if (!string.IsNullOrEmpty(pdfDir))
            Directory.CreateDirectory(pdfDir);

        if (File.Exists(pdfPath))
            File.Delete(pdfPath);

        File.Move(tempPdfPath, pdfPath);
    }

    public static InvalidOperationException WrapExportException(string step, Exception ex)
    {
        var message = new StringBuilder();
        message.Append($"Excel PDF 匯出失敗（步驟：{step}）。");

        if (ex is COMException com)
            message.Append($" HRESULT: 0x{com.HResult:X8}。");

        message.Append($" {ex.Message}");

        if (ex.InnerException is not null)
            message.Append($" InnerException: {ex.InnerException.Message}");

        return new InvalidOperationException(message.ToString(), ex);
    }

    private static void ExportWorksheetAsPdf(dynamic worksheet, string pdfPath)
    {
        var normalizedPath = Path.GetFullPath(pdfPath);

        worksheet.ExportAsFixedFormat(
            XlTypePdf,
            normalizedPath,
            Type.Missing,
            true,
            false,
            Type.Missing,
            Type.Missing,
            false,
            Type.Missing);
    }

    private static (dynamic ExportWorkbook, bool CloseWhenDone, string? TempXlsxPath) OpenSaveCopy(
        dynamic excel,
        dynamic sourceWorkbook,
        string sourcePath)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "OfficeLegacyConverter", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);

        var ext = Path.GetExtension(sourcePath);
        if (string.IsNullOrEmpty(ext))
            ext = ".xlsx";

        var tempXlsx = Path.Combine(tempDir, "export" + ext);

        try
        {
            sourceWorkbook.SaveCopyAs(tempXlsx);
        }
        catch (Exception ex)
        {
            TryDeleteDirectory(tempDir);
            throw WrapExportException("SaveCopyAs 至暫存路徑", ex);
        }

        try
        {
            var copy = excel.Workbooks.Open(
                tempXlsx,
                0,
                false,
                Type.Missing,
                Type.Missing,
                Type.Missing,
                true);
            return (copy, true, tempXlsx);
        }
        catch (Exception ex)
        {
            TryDeleteDirectory(tempDir);
            throw WrapExportException("開啟 SaveCopyAs 副本", ex);
        }
    }

    private static string SanitizeFileName(string fileName)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var sb = new StringBuilder(fileName.Length);
        foreach (var c in fileName)
        {
            if (c == '&' || c == '%' || c == '#' || Array.IndexOf(invalid, c) >= 0)
                sb.Append('_');
            else
                sb.Append(c);
        }

        var sanitized = sb.ToString();
        return string.IsNullOrWhiteSpace(sanitized) ? "workbook.xlsx" : sanitized;
    }

    private static string CopyToTempSafe(string inputPath)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "OfficeLegacyConverter", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempDir);
        var tempCopy = Path.Combine(tempDir, SanitizeFileName(Path.GetFileName(inputPath)));
        File.Copy(inputPath, tempCopy, overwrite: true);
        return tempCopy;
    }

    private static bool IsFileAccessComError(COMException ex) =>
        (uint)ex.HResult == 0x800A03EC || (uint)ex.HResult == 0x80010001;

    public static void TryDeleteFile(string? path)
    {
        if (string.IsNullOrEmpty(path))
            return;

        TryDeleteDirectory(Path.GetDirectoryName(path));
    }

    private static void TryDeleteDirectory(string? path)
    {
        if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
            return;

        try
        {
            Directory.Delete(path, recursive: true);
        }
        catch
        {
            // ignore
        }
    }
}
