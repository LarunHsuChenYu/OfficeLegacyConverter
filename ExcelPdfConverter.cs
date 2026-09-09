using System.Runtime.InteropServices;

namespace OfficeLegacyConverter;

internal static class ExcelPdfConverter
{
    public static void ConvertFile(
        string inputPath,
        string pdfPath,
        bool showExcel,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        dynamic? excel = null;
        var ownsExcelInstance = false;
        dynamic? workbook = null;
        var openedByUs = false;
        string? tempInputCopy = null;
        string? tempPdfPath = null;
        dynamic? exportWorkbook = null;
        var closeExportWorkbook = false;
        string? tempExportXlsx = null;
        List<(dynamic Sheet, int OriginalVisible)>? hiddenBlankSheets = null;

        try
        {
            excel = ExcelComHelper.CreateDedicatedExcelForPdfExport();
            ownsExcelInstance = true;
            excel.Visible = showExcel;
            ExcelComHelper.ConfigureExcelForExport(excel);

            cancellationToken.ThrowIfCancellationRequested();
            progress.Report((current, total, $"開啟 {inputPath}..."));

            (workbook, openedByUs, tempInputCopy) = ExcelComHelper.OpenWorkbookForPdfExport((object)excel!, inputPath);
            var sheetCount = (int)workbook.Worksheets.Count;

            // 在執行任何匯出策略前，先對來源工作簿套用列印版面並隱藏空白表（SaveCopyAs 副本會繼承此狀態）。
            hiddenBlankSheets = ExcelComHelper.ApplyPrintLayout(
                excel!,
                workbook,
                sheetCount,
                progress,
                current,
                total,
                cancellationToken);

            tempPdfPath = ExcelComHelper.CreateSafeTempPdfPath();

            cancellationToken.ThrowIfCancellationRequested();
            progress.Report((current, total, $"匯出 PDF（{sheetCount} 個工作表）..."));

            ExportWorkbookToPdf(
                excel!,
                workbook,
                inputPath,
                tempPdfPath,
                sheetCount,
                current,
                total,
                progress,
                cancellationToken,
                out exportWorkbook,
                out closeExportWorkbook,
                out tempExportXlsx);

            ExcelComHelper.PublishTempPdf(tempPdfPath, pdfPath);
            tempPdfPath = null;

            progress.Report((current, total, $"已匯出 {sheetCount} 個工作表至 {pdfPath}"));
            progress.Report((current, total, $"完成：{Path.GetFileName(inputPath)} → {Path.GetFileName(pdfPath)}"));
        }
        catch (Exception ex) when (ex is not InvalidOperationException)
        {
            throw ExcelComHelper.WrapExportException("Excel 轉 PDF", ex);
        }
        finally
        {
            if (closeExportWorkbook && exportWorkbook is not null)
            {
                try { exportWorkbook.Close(false); } catch { /* ignore */ }
            }

            // 還原暫時隱藏的空白工作表（須在來源工作簿 Close 之前）。
            if (hiddenBlankSheets is not null)
                ExcelComHelper.RestoreSheetVisibility(hiddenBlankSheets);

            if (openedByUs && workbook is not null)
            {
                try { workbook.Close(false); } catch { /* ignore */ }
            }

            ExcelComHelper.TryDeleteFile(tempPdfPath);
            ExcelComHelper.TryDeleteFile(tempInputCopy);
            ExcelComHelper.TryDeleteFile(tempExportXlsx);

            if (excel is not null)
            {
                if (ownsExcelInstance)
                {
                    try { excel.Quit(); } catch { /* ignore */ }
                }

                Marshal.ReleaseComObject(excel);
            }
        }
    }

    private static void ExportWorkbookToPdf(
        dynamic excel,
        dynamic workbook,
        string inputPath,
        string tempPdfPath,
        int sheetCount,
        int current,
        int total,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken,
        out dynamic? exportWorkbook,
        out bool closeExportWorkbook,
        out string? tempExportXlsx)
    {
        exportWorkbook = null;
        closeExportWorkbook = false;
        tempExportXlsx = null;

        var strategies = BuildExportStrategies(sheetCount);

        Exception? lastError = null;

        foreach (var strategy in strategies)
        {
            cancellationToken.ThrowIfCancellationRequested();

            dynamic? activeWorkbook = null;
            var closeActive = false;
            string? activeTempXlsx = null;

            try
            {
                progress.Report((current, total, strategy.ProgressMessage));

                activeWorkbook = ResolveWorkbookForStrategy(
                    excel,
                    workbook,
                    inputPath,
                    strategy,
                    ref exportWorkbook,
                    ref closeExportWorkbook,
                    ref tempExportXlsx,
                    out closeActive,
                    out activeTempXlsx);

                strategy.Export(
                    activeWorkbook,
                    tempPdfPath,
                    sheetCount,
                    progress,
                    current,
                    total,
                    cancellationToken);

                if (!File.Exists(tempPdfPath) || new FileInfo(tempPdfPath).Length == 0)
                    throw new InvalidOperationException($"策略「{strategy.Name}」未產生有效 PDF：{tempPdfPath}");

                return;
            }
            catch (Exception ex)
            {
                lastError = ex;
                TryDeletePartialPdf(tempPdfPath);
            }
            finally
            {
                if (closeActive && activeWorkbook is not null)
                {
                    try { activeWorkbook.Close(false); } catch { /* ignore */ }

                    if (ReferenceEquals(activeWorkbook, exportWorkbook))
                    {
                        exportWorkbook = null;
                        closeExportWorkbook = false;
                        tempExportXlsx = null;
                    }
                }

                if (closeActive && activeTempXlsx is not null)
                    ExcelComHelper.TryDeleteFile(activeTempXlsx);
            }
        }

        throw lastError is not null
            ? ExcelComHelper.WrapExportException("所有匯出策略均失敗", lastError)
            : new InvalidOperationException("Excel PDF 匯出失敗：沒有可用的匯出策略。");
    }

    private static IReadOnlyList<ExportStrategy> BuildExportStrategies(int sheetCount)
    {
        var strategies = new List<ExportStrategy>
        {
            new(
                "整本工作簿直接匯出",
                "匯出整本工作簿為 PDF...",
                useSaveCopy: false,
                perSheet: false),
            new(
                "SaveCopyAs 後整本匯出",
                "建立暫存副本後匯出 PDF...",
                useSaveCopy: true,
                perSheet: false),
        };

        if (sheetCount > 1)
        {
            strategies.Add(new ExportStrategy(
                "逐工作表匯出後合併",
                $"逐工作表匯出（共 {sheetCount} 個）...",
                useSaveCopy: false,
                perSheet: true));

            strategies.Add(new ExportStrategy(
                "SaveCopyAs 後逐工作表匯出",
                $"建立暫存副本後逐工作表匯出（共 {sheetCount} 個）...",
                useSaveCopy: true,
                perSheet: true));
        }

        return strategies;
    }

    private static dynamic ResolveWorkbookForStrategy(
        dynamic excel,
        dynamic sourceWorkbook,
        string inputPath,
        ExportStrategy strategy,
        ref dynamic? exportWorkbook,
        ref bool closeExportWorkbook,
        ref string? tempExportXlsx,
        out bool closeActive,
        out string? activeTempXlsx)
    {
        closeActive = false;
        activeTempXlsx = null;

        if (!strategy.UseSaveCopy)
            return sourceWorkbook;

        if (exportWorkbook is not null)
            return exportWorkbook;

        var prepared = ExcelComHelper.PrepareWorkbookForExport(
            excel,
            sourceWorkbook,
            inputPath,
            forceSaveCopy: true);
        exportWorkbook = prepared.Item1;
        closeExportWorkbook = prepared.Item2;
        tempExportXlsx = prepared.Item3;

        if (closeExportWorkbook)
        {
            closeActive = true;
            activeTempXlsx = tempExportXlsx;
        }

        return exportWorkbook!;
    }

    private static void TryDeletePartialPdf(string tempPdfPath)
    {
        try
        {
            if (File.Exists(tempPdfPath))
                File.Delete(tempPdfPath);
        }
        catch
        {
            // ignore
        }
    }

    private sealed class ExportStrategy
    {
        public ExportStrategy(string name, string progressMessage, bool useSaveCopy, bool perSheet)
        {
            Name = name;
            ProgressMessage = progressMessage;
            UseSaveCopy = useSaveCopy;
            PerSheet = perSheet;
        }

        public string Name { get; }
        public string ProgressMessage { get; }
        public bool UseSaveCopy { get; }
        public bool PerSheet { get; }

        public void Export(
            dynamic workbook,
            string tempPdfPath,
            int sheetCount,
            IProgress<(int current, int total, string message)> progress,
            int current,
            int total,
            CancellationToken cancellationToken)
        {
            if (PerSheet)
            {
                ExcelComHelper.ExportSheetsAsPdfAndMerge(
                    workbook,
                    sheetCount,
                    tempPdfPath,
                    progress,
                    current,
                    total,
                    cancellationToken);
                return;
            }

            try
            {
                ExcelComHelper.ExportWorkbookAsPdf(workbook, tempPdfPath);
            }
            catch (Exception ex)
            {
                throw ExcelComHelper.WrapExportException(Name, ex);
            }
        }
    }
}
