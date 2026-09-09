using System.Text.Json;
using OfficeLegacyConverter;

var outDir = args.Length > 0 ? args[0] : @"d:\VS\Doc2Docx\_regress_kpi";
var requestedPdf = args.Length > 1 ? args[1] : null;
var pdf = requestedPdf ?? Directory
    .GetFiles(@"d:\07-DGPM\DGPM_SPM\docs\requirements", "*KPI*.pdf", SearchOption.AllDirectories)
    .Where(f => !f.Contains($"{Path.DirectorySeparatorChar}Mail{Path.DirectorySeparatorChar}", StringComparison.OrdinalIgnoreCase))
    .OrderByDescending(File.GetLastWriteTimeUtc)
    .FirstOrDefault();
if (pdf is null)
{
    Console.Error.WriteLine("PDF not found");
    return 1;
}

var baseName = Path.GetFileNameWithoutExtension(pdf);
var rawPath = Path.Combine(outDir, baseName + ".raw.md");
if (!File.Exists(rawPath))
{
    Console.Error.WriteLine($"raw.md missing: {rawPath}");
    return 1;
}

Directory.SetCurrentDirectory(Path.GetDirectoryName(typeof(PdfEnhancementService).Assembly.Location)!);
MarkItDownConverter.CheckEnvironment();
Console.WriteLine(MarkItDownConverter.EnvironmentStatus);
if (!MarkItDownConverter.IsEnvironmentReady)
    return 1;

var enhance = PdfEnhancementService.Enhance(pdf, outDir, baseName);
Console.WriteLine($"enhance pages={enhance.PagesExported} total={enhance.TablesExtractedTotal} good={enhance.TablesExtractedGood} err={enhance.Error}");
Console.WriteLine($"likely={string.Join(",", enhance.Pages.Where(p => p.LikelyImageTable).Select(p => p.PageNumber))}");
var expectedMissingPages = Enumerable.Range(4, Math.Max(0, enhance.PagesExported - 3)).ToArray();
var actualLikelyPages = enhance.Pages.Where(p => p.LikelyImageTable).Select(p => p.PageNumber).ToArray();
Console.WriteLine($"p1 likely={enhance.Pages.Single(p => p.PageNumber == 1).LikelyImageTable}");
Console.WriteLine($"ocr ok={enhance.Pages.Count(p => p.OcrStatus == "ok")}");

var result = MarkdownPostProcessor.Process(
    rawPath,
    pdf,
    outDir,
    enhance,
    conversionEngine: "MarkItDown",
    keepRawBackup: true,
    writeQualityJson: true);

var qualityPath = result.QualityJsonPath!;
var json = JsonDocument.Parse(File.ReadAllText(qualityPath));
var gaps = json.RootElement.GetProperty("gaps").EnumerateArray()
    .Select(g => (Type: g.GetProperty("Type").GetString(), Loc: g.GetProperty("Location").GetString()))
    .ToList();
var missing = gaps.Where(g => g.Type == "missing_table").ToList();
Console.WriteLine($"missing_table count={missing.Count}");
foreach (var g in missing)
    Console.WriteLine($"  {g.Loc}");

var md = File.ReadAllText(result.FinalMdPath);
var lossMarkers = System.Text.RegularExpressions.Regex.Matches(md, @"\[表格遺失：影像表 p\.(\d+)\]");
Console.WriteLine($"md missing markers={lossMarkers.Count}");
Console.WriteLine($"md has 圖片描述={(md.Contains("圖片描述：", StringComparison.Ordinal) ? "yes" : "no")}");
Console.WriteLine($"md has 圖片索引={(md.Contains("## 圖片索引", StringComparison.Ordinal) ? "yes" : "no")}");
Console.WriteLine($"raw has 圖片索引={(File.ReadAllText(result.RawMdPath!).Contains("## 圖片索引", StringComparison.Ordinal) ? "yes" : "no")}");

var semanticCoverageIsDynamic =
    md.Contains($"偵測到 {expectedMissingPages.Length} 頁影像表格未轉為可搜尋文字", StringComparison.Ordinal)
    && !md.Contains("通常 90%+", StringComparison.Ordinal);
var ocrPages = enhance.Pages.Where(p => expectedMissingPages.Contains(p.PageNumber)).ToList();
var page4Ocr = enhance.Pages.Single(p => p.PageNumber == 4).OcrPreview;
var sampleValueFound = enhance.PagesExported != 16
    || page4Ocr.Contains("15,644,421", StringComparison.Ordinal);
Console.WriteLine($"p4 sample value found={sampleValueFound}");
var ok =
    actualLikelyPages.SequenceEqual(expectedMissingPages)
    && missing.Count == expectedMissingPages.Length
    && lossMarkers.Count == expectedMissingPages.Length
    && ocrPages.All(p => p.OcrStatus == "ok" && !string.IsNullOrWhiteSpace(p.OcrPreview))
    && sampleValueFound
    && semanticCoverageIsDynamic;
Console.WriteLine(ok ? "REGRESS_OK" : "REGRESS_FAIL");
return ok ? 0 : 2;
