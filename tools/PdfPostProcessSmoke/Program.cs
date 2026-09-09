using System.Text.Json;
using System.Text.RegularExpressions;
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
    .Select(g => (
        Type: g.GetProperty("Type").GetString(),
        Loc: g.GetProperty("Location").GetString(),
        Recovery: g.TryGetProperty("recovery_status", out var recovery)
            ? recovery.GetString()
            : null))
    .ToList();
var missing = gaps.Where(g => g.Type == "missing_table").ToList();
var qualityPages = json.RootElement
    .GetProperty("pdf_enhancement")
    .GetProperty("pages")
    .EnumerateArray()
    .Where(p => expectedMissingPages.Contains(p.GetProperty("page").GetInt32()))
    .ToList();
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
    md.Contains($"{expectedMissingPages.Length} 頁影像表格未還原為結構化表格；已有 OCR 可搜尋文字", StringComparison.Ordinal)
    && !md.Contains("通常 90%+", StringComparison.Ordinal);
var qualityReportsCompleteOcr = qualityPages.All(p =>
    !p.GetProperty("ocr_truncated").GetBoolean()
    && p.GetProperty("ocr_text_length").GetInt32()
        == p.GetProperty("ocr_preview_length").GetInt32());
var ocrPages = enhance.Pages.Where(p => expectedMissingPages.Contains(p.PageNumber)).ToList();
var page4Ocr = enhance.Pages.Single(p => p.PageNumber == 4).OcrPreview;
var sampleValueFound = enhance.PagesExported != 16
    || page4Ocr.Contains("15,644,421", StringComparison.Ordinal);
Console.WriteLine($"p4 sample value found={sampleValueFound}");
var fullTextNotTruncated = enhance.PagesExported != 16
    || new[] { 4, 7, 10 }.All(page =>
    {
        var text = enhance.Pages.Single(p => p.PageNumber == page).OcrPreview;
        return text.Length > 1501 && !text.EndsWith('…');
    });
var page4TailFound = enhance.PagesExported != 16
    || page4Ocr.Contains("Key Reads", StringComparison.OrdinalIgnoreCase);
var page7Ocr = enhance.Pages.Single(p => p.PageNumber == 7).OcrPreview;
var page7TailFound = enhance.PagesExported != 16
    || new[] { "675", "488", "77", "29", "159", "107", "34", "8" }
        .All(value => page7Ocr.Contains(value, StringComparison.Ordinal));
var page9Ocr = enhance.Pages.Single(p => p.PageNumber == 9).OcrPreview;
var page9RiValuesFound = enhance.PagesExported != 16
    || Regex.IsMatch(
        page9Ocr,
        @"RI Sales / Penetration Carpark\s*-\s*%.*?\b13\b.*?\b2\b",
        RegexOptions.IgnoreCase | RegexOptions.Singleline);
var page10Ocr = enhance.Pages.Single(p => p.PageNumber == 10).OcrPreview;
var page10TailFound = enhance.PagesExported != 16
    || new[] { "20,811", "2,671", "3,999" }
        .All(value => page10Ocr.Contains(value, StringComparison.Ordinal));
var page14Ocr = enhance.Pages.Single(p => p.PageNumber == 14).OcrPreview;
var page14ValuesFound = enhance.PagesExported != 16
    || new[] { "49,566,416", "30,099,945" }
        .All(value => page14Ocr.Contains(value, StringComparison.Ordinal));
var noSparseSupplementNoise =
    !md.Contains("OCR 補充辨識（稀疏數值）", StringComparison.Ordinal);
var recoveryStatusIsExplicit = missing.All(
    gap => gap.Recovery == "ocr_available_unverified");
Console.WriteLine($"full OCR not truncated={fullTextNotTruncated}");
Console.WriteLine($"p4 tail found={page4TailFound}");
Console.WriteLine($"p7 tail found={page7TailFound}");
Console.WriteLine($"p9 RI values found={page9RiValuesFound}");
Console.WriteLine($"p10 tail found={page10TailFound}");
Console.WriteLine($"p14 values found={page14ValuesFound}");
Console.WriteLine($"no sparse supplement noise={noSparseSupplementNoise}");
Console.WriteLine($"recovery status explicit={recoveryStatusIsExplicit}");
Console.WriteLine($"quality reports complete OCR={qualityReportsCompleteOcr}");
var ok =
    actualLikelyPages.SequenceEqual(expectedMissingPages)
    && missing.Count == expectedMissingPages.Length
    && lossMarkers.Count == expectedMissingPages.Length
    && ocrPages.All(p => p.OcrStatus == "ok" && !string.IsNullOrWhiteSpace(p.OcrPreview))
    && sampleValueFound
    && fullTextNotTruncated
    && page4TailFound
    && page7TailFound
    && page9RiValuesFound
    && page10TailFound
    && page14ValuesFound
    && noSparseSupplementNoise
    && recoveryStatusIsExplicit
    && qualityReportsCompleteOcr
    && semanticCoverageIsDynamic;
Console.WriteLine(ok ? "REGRESS_OK" : "REGRESS_FAIL");
return ok ? 0 : 2;
