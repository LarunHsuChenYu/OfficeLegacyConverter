using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeLegacyConverter;

internal sealed class XlsxConversionResult
{
    public string Markdown { get; init; } = "";
    public int SheetCount { get; init; }
    public IReadOnlyList<string> SheetNames { get; init; } = [];
    public IReadOnlyList<string> ConversionLog { get; init; } = [];
}

internal static class XlsxMarkdownConverter
{
    private static readonly Regex CellReferencePattern = new(
        @"^(?<col>[A-Z]+)(?<row>\d+)$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static XlsxConversionResult Convert(
        string xlsxPath,
        Action<string>? logSheetProgress = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var sb = new StringBuilder();
        var conversionLog = new List<string>();
        var sheetNames = new List<string>();

        using var document = SpreadsheetDocument.Open(xlsxPath, false);
        var workbookPart = document.WorkbookPart
            ?? throw new InvalidOperationException("無法讀取 Excel 活頁簿。");

        var sharedStrings = BuildSharedStringTable(workbookPart);
        var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
        var totalSheets = sheets.Count;

        var sheetIndex = 0;
        foreach (var sheet in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheetIndex++;

            if (sheet.Id?.Value is not { } relId)
            {
                conversionLog.Add($"略過無效工作表項目（{sheetIndex}/{totalSheets}）");
                continue;
            }

            if (workbookPart.GetPartById(relId) is not WorksheetPart worksheetPart)
            {
                conversionLog.Add($"略過無法讀取的工作表（{sheetIndex}/{totalSheets}）");
                continue;
            }

            var sheetName = sheet.Name?.Value ?? $"Sheet{sheetIndex}";
            sheetNames.Add(sheetName);
            logSheetProgress?.Invoke($"已轉換工作表：{sheetName} ({sheetIndex}/{totalSheets})");
            conversionLog.Add($"已轉換工作表：{sheetName} ({sheetIndex}/{totalSheets})");

            sb.AppendLine($"## {sheetName}");
            sb.AppendLine();

            var rows = ReadSheetRows(worksheetPart, sharedStrings);
            if (rows.Count == 0)
            {
                sb.AppendLine("（空白工作表）");
                sb.AppendLine();
                continue;
            }

            var columnCount = rows.Max(r => r.Count);
            for (var r = 0; r < rows.Count; r++)
            {
                var row = rows[r];
                while (row.Count < columnCount)
                    row.Add("");

                sb.Append("| ");
                sb.Append(string.Join(" | ", row.Select(EscapeTableCell)));
                sb.AppendLine(" |");

                if (r == 0)
                {
                    sb.Append("| ");
                    sb.Append(string.Join(" | ", Enumerable.Repeat("---", columnCount)));
                    sb.AppendLine(" |");
                }
            }

            sb.AppendLine();
        }

        return new XlsxConversionResult
        {
            Markdown = sb.ToString().TrimEnd() + Environment.NewLine,
            SheetCount = sheetNames.Count,
            SheetNames = sheetNames,
            ConversionLog = conversionLog
        };
    }

    private static List<List<string>> ReadSheetRows(WorksheetPart worksheetPart, IReadOnlyList<string> sharedStrings)
    {
        var sheetData = worksheetPart.Worksheet?.Elements<SheetData>().FirstOrDefault();
        var sparseRows = new SortedDictionary<int, Dictionary<int, string>>();

        if (sheetData is not null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                var rowIndex = row.RowIndex is not null ? (int)row.RowIndex.Value : sparseRows.Count + 1;
                var cells = sparseRows.TryGetValue(rowIndex, out var existing)
                    ? existing
                    : sparseRows[rowIndex] = new Dictionary<int, string>();

                foreach (var cell in row.Elements<Cell>())
                {
                    var columnIndex = GetColumnIndex(cell.CellReference?.Value);
                    cells[columnIndex] = GetCellValue(cell, sharedStrings);
                }
            }
        }

        var dimension = ParseSheetDimension(worksheetPart.Worksheet?.SheetDimension?.Reference?.Value);
        if (dimension is not null)
        {
            for (var r = dimension.Value.MinRow; r <= dimension.Value.MaxRow; r++)
            {
                if (!sparseRows.ContainsKey(r))
                    sparseRows[r] = new Dictionary<int, string>();
            }
        }

        if (sparseRows.Count == 0)
            return [];

        var maxRow = dimension?.MaxRow ?? sparseRows.Keys.Max();
        var minRow = dimension?.MinRow ?? sparseRows.Keys.Min();
        var maxCol = Math.Max(
            dimension?.MaxCol ?? 0,
            sparseRows.Values.SelectMany(c => c.Keys).DefaultIfEmpty(0).Max());
        var minCol = dimension?.MinCol ?? 0;
        var result = new List<List<string>>();

        for (var r = minRow; r <= maxRow; r++)
        {
            var rowCells = new List<string>();
            sparseRows.TryGetValue(r, out var sparseCells);
            sparseCells ??= new Dictionary<int, string>();

            for (var c = minCol; c <= maxCol; c++)
                rowCells.Add(sparseCells.TryGetValue(c, out var value) ? value : "");

            if (rowCells.Any(cell => !string.IsNullOrWhiteSpace(cell)))
                result.Add(rowCells);
        }

        return result;
    }

    private static (int MinRow, int MaxRow, int MinCol, int MaxCol)? ParseSheetDimension(string? reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return null;

        var parts = reference.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
            return null;

        var start = ParseCellReference(parts[0]);
        var end = parts.Length > 1 ? ParseCellReference(parts[1]) : start;
        if (start is null || end is null)
            return null;

        return (
            Math.Min(start.Value.Row, end.Value.Row),
            Math.Max(start.Value.Row, end.Value.Row),
            Math.Min(start.Value.Col, end.Value.Col),
            Math.Max(start.Value.Col, end.Value.Col));
    }

    private static (int Row, int Col)? ParseCellReference(string reference)
    {
        var match = CellReferencePattern.Match(reference.ToUpperInvariant());
        if (!match.Success)
            return null;

        if (!int.TryParse(match.Groups["row"].Value, out var row))
            return null;

        return (row, GetColumnIndex(match.Groups["col"].Value));
    }

    private static IReadOnlyList<string> BuildSharedStringTable(WorkbookPart workbookPart)
    {
        if (workbookPart.SharedStringTablePart?.SharedStringTable is not { } table)
            return [];

        return table.Elements<SharedStringItem>()
            .Select(item => string.Concat(item.Descendants<Text>().Select(t => t.Text)))
            .ToList();
    }

    private static string GetCellValue(Cell cell, IReadOnlyList<string> sharedStrings)
    {
        if (cell.CellValue is null)
            return "";

        var raw = cell.CellValue.Text ?? "";
        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(raw, out var index)
            && index >= 0
            && index < sharedStrings.Count)
            return sharedStrings[index];

        if (cell.DataType?.Value == CellValues.Boolean)
            return raw == "1" ? "TRUE" : "FALSE";

        return raw;
    }

    private static int GetColumnIndex(string? cellReference)
    {
        if (string.IsNullOrEmpty(cellReference))
            return 0;

        var letters = new string(cellReference.TakeWhile(char.IsLetter).ToArray()).ToUpperInvariant();
        var index = 0;
        foreach (var ch in letters)
            index = index * 26 + (ch - 'A' + 1);
        return Math.Max(0, index - 1);
    }

    private static string EscapeTableCell(string text) =>
        text.Replace("|", "\\|", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal).Trim();
}
