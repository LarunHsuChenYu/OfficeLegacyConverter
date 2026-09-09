namespace OfficeLegacyConverter;

/// <summary>
/// 單一入口批次升級舊版 Office：.doc→.docx、.xls→.xlsx、.ppt→.pptx。
/// 依副檔名分組後沿用既有 Word／Excel／PowerPoint COM 轉換器。
/// </summary>
internal static class LegacyOfficeUpgradePipeline
{
    public static readonly string[] SupportedExtensions = [".doc", ".xls", ".ppt"];

    private static readonly string LegacyFileFilter =
        "舊版 Office|*.doc;*.xls;*.ppt|Word 97-2003 (*.doc)|*.doc|Excel 97-2003 (*.xls)|*.xls|PowerPoint 97-2003 (*.ppt)|*.ppt";

    public static string FileDialogFilter => LegacyFileFilter;

    public static bool IsSupported(string path) =>
        SupportedExtensions.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase);

    public static IEnumerable<string> ExpandPaths(string path)
    {
        if (!Directory.Exists(path))
            return IsSupported(path) ? [path] : [];

        return SupportedExtensions
            .SelectMany(ext => Directory.EnumerateFiles(path, $"*{ext}", SearchOption.AllDirectories))
            .Where(IsSupported);
    }

    public static BatchResult ConvertFiles(
        IReadOnlyList<string> paths,
        bool replace,
        bool showOffice,
        IProgress<(int current, int total, string message)> progress,
        CancellationToken cancellationToken)
    {
        var total = paths.Count;
        var result = new BatchResult { Total = total };

        var docs = new List<string>();
        var xlss = new List<string>();
        var ppts = new List<string>();
        var unsupported = new List<(int index, string path)>();

        for (var i = 0; i < paths.Count; i++)
        {
            var path = paths[i];
            var ext = Path.GetExtension(path);
            if (ext.Equals(".doc", StringComparison.OrdinalIgnoreCase))
                docs.Add(path);
            else if (ext.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                xlss.Add(path);
            else if (ext.Equals(".ppt", StringComparison.OrdinalIgnoreCase))
                ppts.Add(path);
            else
                unsupported.Add((i, path));
        }

        foreach (var (index, path) in unsupported)
        {
            var current = index + 1;
            var message = $"不支援的副檔名：{Path.GetFileName(path)}";
            result.Failures.Add($"{Path.GetFileName(path)}：僅支援 .doc / .xls / .ppt");
            progress.Report((current, total, $"錯誤 - {path}：{message}"));
        }

        var offset = 0;
        RunGroup(
            docs,
            (group, groupProgress, token) => WordConverter.ConvertFiles(group, replace, showOffice, groupProgress, token),
            ref offset,
            total,
            progress,
            result,
            cancellationToken);

        RunGroup(
            xlss,
            (group, groupProgress, token) => ExcelConverter.ConvertFiles(group, replace, showOffice, groupProgress, token),
            ref offset,
            total,
            progress,
            result,
            cancellationToken);

        RunGroup(
            ppts,
            (group, groupProgress, token) => PowerPointConverter.ConvertFiles(group, replace, showOffice, groupProgress, token),
            ref offset,
            total,
            progress,
            result,
            cancellationToken);

        return result;
    }

    private static void RunGroup(
        IReadOnlyList<string> group,
        Func<IReadOnlyList<string>, IProgress<(int current, int total, string message)>, CancellationToken, BatchResult> convert,
        ref int offset,
        int overallTotal,
        IProgress<(int current, int total, string message)> progress,
        BatchResult aggregate,
        CancellationToken cancellationToken)
    {
        if (group.Count == 0)
            return;

        var groupOffset = offset;
        var groupProgress = new Progress<(int current, int total, string message)>(report =>
        {
            var overallCurrent = Math.Min(groupOffset + report.current, overallTotal);
            progress.Report((overallCurrent, overallTotal, report.message));
        });

        var groupResult = convert(group, groupProgress, cancellationToken);
        aggregate.Failures.AddRange(groupResult.Failures);
        offset += group.Count;
    }
}
