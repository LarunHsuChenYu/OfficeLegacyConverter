using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeLegacyConverter;

internal static class MarkdownPostProcessor
{
    private static readonly Regex StandalonePageNumber = new(@"^\d{1,2}$", RegexOptions.Compiled);
    private static readonly Regex ReferenceMaterialPage = new(@"參考資料：\r?\n(\d{1,2})\r?\n", RegexOptions.Compiled);
    private static readonly Regex ReferenceImagePlaceholder = new(
        @"參考資料：\r?\n> \*\*\[圖片補足：p\.(\d{1,2})(?:\s+見原始 PDF)?\]\*\*\r?\n(?:> .*\r?\n)*",
        RegexOptions.Compiled);
    private static readonly Regex NumberedDataSourceLine = new(@"\(3\)\s*資料來源：", RegexOptions.Compiled);
    private static readonly Regex UnnumberedDataSourceLine = new(@"^\s*資料來源：", RegexOptions.Compiled);
    private static readonly Regex RequirementHeader = new(@"\(4\)\s*需求說明：\s*$", RegexOptions.Compiled);
    private static readonly Regex UnnumberedRequirementHeader = new(@"^\s*需求說明：", RegexOptions.Compiled);
    private static readonly Regex FunctionalScreenHeader = new(@"^\s*功能畫面：", RegexOptions.Compiled);
    private static readonly Regex RequirementAOnly = new(@"^A\.\s*$", RegexOptions.Compiled);
    private static readonly Regex StructuredFieldMarker = new(@"^\(\d+\)\s", RegexOptions.Compiled);
    private static readonly Regex SectionHeader = new(@"^\d+(?:\.\d+)+", RegexOptions.Compiled);
    private static readonly Regex BracketSectionRef = new(@"\[\d+(?:\.\d+)+\]", RegexOptions.Compiled);
    private static readonly Regex MarkdownHeadingSection = new(@"^#+\s*\d+(?:\.\d+)+", RegexOptions.Compiled);
    private static readonly Regex SectionRefPattern = new(
        @"\d+\.\d+\.\[\d+\.\d+\][^\n|｜\r]*",
        RegexOptions.Compiled);
    private static readonly Regex PageFractionMarker = new(
        @"\((\d+)\s*/\s*(\d+)\)",
        RegexOptions.Compiled);

    private static string InjectYamlFrontMatter(string mdContent, MarkdownMetadata metadata)
    {
        var body = StripExistingFrontMatter(mdContent);
        var sb = new StringBuilder();
        sb.AppendLine("---");
        sb.AppendLine($"source_file: {YamlQuote(metadata.SourceFile)}");
        sb.AppendLine($"source_path: {YamlQuote(metadata.SourcePath.Replace('\\', '/'))}");
        sb.AppendLine($"converted_file: {YamlQuote(metadata.ConvertedFile)}");
        sb.AppendLine($"converted_date: {metadata.ConvertedDate:yyyy-MM-dd}");
        sb.AppendLine($"converter: {YamlQuote(metadata.Converter)}");
        sb.AppendLine($"conversion_engine: {YamlQuote(metadata.ConversionEngine)}");
        if (!string.IsNullOrEmpty(metadata.ConverterVersion))
            sb.AppendLine($"converter_version: {YamlQuote(metadata.ConverterVersion)}");
        sb.AppendLine();
        sb.AppendLine($"knowledge_tier: {metadata.KnowledgeTier}");
        sb.AppendLine($"document_type: {YamlQuote(metadata.DocumentType)}");
        if (!string.IsNullOrEmpty(metadata.ImagesFolder))
            sb.AppendLine($"images_folder: {YamlQuote(metadata.ImagesFolder)}");
        sb.AppendLine();
        sb.AppendLine("quality:");
        var missingTablePages = metadata.KnownGaps
            .Where(g => g.Type == "missing_table")
            .SelectMany(g => g.PdfPages ?? [])
            .Distinct()
            .Count();
        var coverageMessage = missingTablePages > 0
            ? $"未評估；偵測到 {missingTablePages} 頁影像表格未轉為可搜尋文字，不得套用固定覆蓋率"
            : "自動轉換，語意覆蓋率待人工評估";
        sb.AppendLine($"  semantic_coverage: {YamlQuote(coverageMessage)}");
        if (metadata.KnownGaps.Count == 0)
        {
            sb.AppendLine("  known_gaps: []");
        }
        else
        {
            sb.AppendLine("  known_gaps:");
            foreach (var gap in metadata.KnownGaps)
            {
                sb.AppendLine($"    - type: {YamlQuote(gap.Type)}");
                sb.AppendLine($"      location: {YamlQuote(gap.Location)}");
                if (!string.IsNullOrEmpty(gap.Field))
                    sb.AppendLine($"      field: {YamlQuote(gap.Field!)}");
                if (!string.IsNullOrEmpty(gap.Description))
                    sb.AppendLine($"      description: {YamlQuote(gap.Description!)}");
                if (!string.IsNullOrEmpty(gap.Severity))
                    sb.AppendLine($"      severity: {YamlQuote(gap.Severity!)}");
                if (gap.PdfPages is { Length: > 0 })
                    sb.AppendLine($"      pdf_pages: [{string.Join(", ", gap.PdfPages)}]");
            }
        }
        sb.AppendLine();
        sb.AppendLine("refinement:");
        sb.AppendLine("  status: \"raw\"");
        sb.AppendLine();
        sb.AppendLine("llm_instructions: |");
        sb.AppendLine("  本文件由 Office/PDF 自動轉換而來，屬知識分級 Tier-1 原始擷取。");
        sb.AppendLine("  遇到 [圖片補足]、[待補]、[表格待修]、[表格遺失] 標記時，以標記內容為準，不得腦補。");
        sb.AppendLine("  引用時請註明 source_file 與章節編號。");
        sb.AppendLine("  若表格破碎或遺失，請對照 source_file、頁面截圖與圖片描述／OCR 摘要；孤立頁碼已清理，勿當作需求內容。");
        sb.AppendLine("  圖片索引與缺表文字描述僅存在於最終 .md（不在 .raw.md）。");
        sb.AppendLine("---");
        sb.AppendLine();
        sb.Append(body.TrimStart());
        return sb.ToString();
    }

    private static (string cleaned, List<string> log) CleanPageArtifacts(string mdContent)
    {
        var log = new List<string>();
        var normalized = mdContent.Replace("\r\n", "\n").Replace('\r', '\n');

        normalized = ReferenceMaterialPage.Replace(normalized, match =>
        {
            var page = match.Groups[1].Value;
            log.Add($"參考資料：頁碼 {page} → 圖片補足占位");
            return $"""
                參考資料：
                > **[圖片補足：p.{page} 見原始 PDF]**
                > - 狀態：自動轉換未擷取圖片，請對照 source_file 原圖補充
                > - 置信度：待人工補充

                """;
        });

        var lines = normalized.Split('\n');
        var result = new List<string>(lines.Length);

        for (var i = 0; i < lines.Length; i++)
        {
            var line = lines[i];

            if (IsGhostTableRow(line))
            {
                log.Add($"刪除表格幽靈列：{TruncateForLog(line)}");
                continue;
            }

            if (IsStandalonePageNumberLine(line, lines, i))
            {
                log.Add($"刪除獨立頁碼行：{line.Trim()}");
                continue;
            }

            result.Add(line);
        }

        return (string.Join(Environment.NewLine, result), log);
    }

    private static List<KnownGap> DetectKnownGaps(string mdContent)
    {
        var gaps = new List<KnownGap>();
        var normalized = mdContent.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');

        for (var i = 0; i < lines.Length; i++)
        {
            if (!IsDataSourceMarkerLine(lines[i], out _, out var isNumbered))
                continue;

            if (!IsTrulyBlankDataSource(lines, i))
                continue;

            var pos = LineStartIndex(normalized, i);
            gaps.Add(new KnownGap
            {
                Type = "blank_section",
                Location = FindNearestSection(normalized, pos),
                Field = DataSourceFieldLabel(isNumbered),
                Description = "資料來源段落為空白",
                Severity = "medium"
            });
        }

        for (var i = 0; i < lines.Length; i++)
        {
            if (!RequirementHeader.IsMatch(lines[i].TrimEnd()))
                continue;

            var aLineIdx = FindNextNonEmptyLineIndex(lines, i + 1);
            if (aLineIdx is null || !RequirementAOnly.IsMatch(lines[aLineIdx.Value].Trim()))
                continue;

            if (!IsTrulyBlankRequirementA(lines, aLineIdx.Value))
                continue;

            var pos = LineStartIndex(normalized, i);
            gaps.Add(new KnownGap
            {
                Type = "blank_section",
                Location = FindNearestSection(normalized, pos),
                Field = "(4) 需求說明 A.",
                Description = "需求說明 A. 後無實質內容",
                Severity = "medium"
            });
        }

        foreach (var tableGap in DetectBrokenTables(normalized))
            gaps.Add(tableGap);

        if (ReferenceMaterialPage.IsMatch(normalized) || normalized.Contains("[圖片補足：p.", StringComparison.Ordinal))
        {
            var refIdx = normalized.IndexOf("參考資料：", StringComparison.Ordinal);
            gaps.Add(new KnownGap
            {
                Type = "missing_image",
                Location = FindReferenceImageLocation(normalized, refIdx),
                Description = "參考資料區段可能含未擷取圖片",
                Severity = "high"
            });
        }

        return DeduplicateGaps(gaps);
    }

    private static List<KnownGap> DetectMissingTables(
        string mdContent,
        PdfEnhancementResult? pdfEnhancement)
    {
        var gaps = new List<KnownGap>();
        if (pdfEnhancement is null || pdfEnhancement.Pages.Count == 0)
            return gaps;

        var sections = BuildPageSections(mdContent);
        foreach (var page in pdfEnhancement.Pages)
        {
            if (!page.LikelyImageTable)
                continue;

            if (PageSectionHasValidTable(mdContent, sections, page.PageNumber))
                continue;

            gaps.Add(new KnownGap
            {
                Type = "missing_table",
                Location = $"p.{page.PageNumber}",
                Description = "文字層無明細表（疑似影像／Excel 截圖表格）",
                Severity = "high",
                PdfPages = [page.PageNumber]
            });
        }

        return gaps;
    }

    private static Dictionary<int, (int Start, int End)> BuildPageSections(string mdContent)
    {
        var normalized = mdContent.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var starts = new List<(int Page, int Line)>();

        for (var i = 0; i < lines.Length; i++)
        {
            var match = PageFractionMarker.Match(lines[i]);
            if (!match.Success)
                continue;
            if (!int.TryParse(match.Groups[1].Value, out var pageNum) || pageNum < 1)
                continue;
            starts.Add((pageNum, i));
        }

        var map = new Dictionary<int, (int Start, int End)>();
        for (var i = 0; i < starts.Count; i++)
        {
            var (page, start) = starts[i];
            var end = i + 1 < starts.Count ? starts[i + 1].Line : lines.Length;
            // Keep first span if duplicate page markers appear
            if (!map.ContainsKey(page))
                map[page] = (start, end);
        }

        return map;
    }

    private static bool PageSectionHasValidTable(
        string mdContent,
        Dictionary<int, (int Start, int End)> sections,
        int pageNumber)
    {
        var normalized = mdContent.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');

        int start;
        int end;
        if (sections.TryGetValue(pageNumber, out var span))
        {
            start = span.Start;
            end = span.End;
        }
        else
        {
            // No (n/N) markers: cannot prove a table belongs to this page
            return false;
        }

        var i = start;
        while (i < end && i < lines.Length)
        {
            if (!IsTableRow(lines[i]))
            {
                i++;
                continue;
            }

            var tableLines = new List<string>();
            while (i < end && i < lines.Length && IsTableRow(lines[i]))
            {
                tableLines.Add(lines[i]);
                i++;
            }

            if (tableLines.Count >= 2)
                return true;
        }

        return false;
    }

    private static string BuildContextCaptionFromMd(
        string mdContent,
        Dictionary<int, (int Start, int End)> sections,
        int pageNumber,
        PdfPageInfo page)
    {
        if (!string.IsNullOrWhiteSpace(page.ImageCaption))
            return page.ImageCaption;

        var normalized = mdContent.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var title = "";
        var bullets = new List<string>();

        if (sections.TryGetValue(pageNumber, out var span))
        {
            for (var i = span.Start; i < span.End && i < lines.Length; i++)
            {
                var trimmed = lines[i].Trim();
                if (string.IsNullOrEmpty(trimmed) || IsTableRow(trimmed))
                    continue;
                if (string.IsNullOrEmpty(title) && PageFractionMarker.IsMatch(trimmed))
                {
                    title = trimmed;
                    continue;
                }

                if (trimmed.StartsWith('➢') || trimmed.StartsWith('•') || trimmed.StartsWith('-') ||
                    trimmed.Contains("共計", StringComparison.Ordinal) ||
                    trimmed.Contains("需換算", StringComparison.Ordinal))
                {
                    bullets.Add(trimmed.TrimStart('➢', '•', '-', ' ', '·'));
                }

                if (bullets.Count >= 3)
                    break;
            }
        }
        else if (page.ContextHints.Count > 0)
        {
            title = page.ContextHints[0];
            bullets.AddRange(page.ContextHints.Skip(1).Take(3));
        }

        var parts = new List<string>
        {
            $"第 {pageNumber} 頁影像表格截圖（Excel 篩選列／明細數據）。"
        };
        if (!string.IsNullOrEmpty(title))
        {
            var shortTitle = title.Length > 80 ? title[..80] + "…" : title;
            parts.Add($"標題「{shortTitle}」。");
        }

        if (bullets.Count > 0)
            parts.Add("頁面文字結論：" + string.Join("；", bullets.Select(b => b.Length > 60 ? b[..60] : b)) + "。");

        parts.Add("自動轉換未能還原儲存格，請對照下方截圖。");
        return string.Concat(parts);
    }

    public static PostProcessResult Process(
        string rawMdPath,
        string sourceFilePath,
        string outputDir,
        PdfEnhancementResult? pdfEnhancement = null,
        DocxConversionResult? docxConversion = null,
        XlsxConversionResult? xlsxConversion = null,
        string conversionEngine = "MarkItDown",
        bool keepRawBackup = true,
        bool writeQualityJson = true)
    {
        var rawContent = File.ReadAllText(rawMdPath);
        var (cleaned, cleanupLog) = CleanPageArtifacts(rawContent);
        var gaps = DetectKnownGaps(cleaned);
        gaps.AddRange(DetectMissingTables(cleaned, pdfEnhancement));
        gaps = DeduplicateGaps(gaps);

        var enhancementLog = new List<string>();
        var processed = cleaned;
        processed = InsertBlankDataSourceAnnotations(processed, gaps, enhancementLog);
        processed = ApplyBrokenTableAnnotations(processed, pdfEnhancement, enhancementLog);
        processed = InsertMissingTableFallbacks(processed, pdfEnhancement, gaps, enhancementLog);
        processed = InsertPdfPageImages(processed, pdfEnhancement, enhancementLog);

        var baseName = Path.GetFileNameWithoutExtension(sourceFilePath);
        var isOpenXml = conversionEngine.Equals("OpenXml", StringComparison.OrdinalIgnoreCase);
        var imagesFolder = pdfEnhancement?.ImagesFolder ?? docxConversion?.ImagesFolder;
        var metadata = new MarkdownMetadata
        {
            SourceFile = Path.GetFileName(sourceFilePath),
            SourcePath = Path.GetFullPath(sourceFilePath),
            ConvertedFile = baseName + ".md",
            ConvertedDate = DateTime.Now,
            Converter = isOpenXml ? "OfficeLegacyConverter + OpenXml" : "OfficeLegacyConverter + markitdown",
            ConversionEngine = conversionEngine,
            ConverterVersion = Assembly.GetExecutingAssembly().GetName().Version?.ToString(3),
            KnowledgeTier = 1,
            DocumentType = InferDocumentType(sourceFilePath),
            ImagesFolder = imagesFolder,
            KnownGaps = gaps
        };

        var finalContent = InjectYamlFrontMatter(processed, metadata);
        var finalPath = Path.Combine(outputDir, baseName + ".md");
        File.WriteAllText(finalPath, finalContent, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        string? rawPath = null;
        if (keepRawBackup)
        {
            rawPath = Path.Combine(outputDir, baseName + ".raw.md");
            if (!Path.GetFullPath(rawMdPath).Equals(Path.GetFullPath(rawPath), StringComparison.OrdinalIgnoreCase))
                File.Copy(rawMdPath, rawPath, overwrite: true);
            else
                rawPath = rawMdPath;
        }

        string? qualityPath = null;
        if (writeQualityJson)
        {
            qualityPath = Path.Combine(outputDir, baseName + "_quality.json");
            WriteQualityJson(qualityPath, metadata, cleanupLog, enhancementLog, gaps, pdfEnhancement, docxConversion, xlsxConversion);
        }

        return new PostProcessResult
        {
            FinalMdPath = finalPath,
            RawMdPath = rawPath,
            QualityJsonPath = qualityPath,
            CleanupLog = cleanupLog,
            DetectedGaps = gaps,
            PdfEnhancement = pdfEnhancement
        };
    }

    private static bool IsStandalonePageNumberLine(string line, string[] lines, int index)
    {
        var trimmed = line.Trim();
        if (!StandalonePageNumber.IsMatch(trimmed))
            return false;

        if (!int.TryParse(trimmed, out var n) || n is < 1 or > 99)
            return false;

        if (trimmed.StartsWith('(') || trimmed.Contains('.'))
            return false;

        if (Regex.IsMatch(trimmed, @"^[A-Za-z]\.$"))
            return false;

        if (Regex.IsMatch(trimmed, @"^\d{6,}："))
            return false;

        if (IsTableRow(line))
            return false;

        var prev = index > 0 ? lines[index - 1].Trim() : "";
        var next = index < lines.Length - 1 ? lines[index + 1].Trim() : "";

        if (Regex.IsMatch(prev, @"^\d+\.\s") || Regex.IsMatch(next, @"^\d+\.\s"))
            return false;

        if (StructuredFieldMarker.IsMatch(prev) || StructuredFieldMarker.IsMatch(next))
            return false;

        if (IsOrphanPageNumberAfterReferenceBlock(lines, index))
            return true;

        // 保留日期脈絡中的年月數字（如「資料年月： 202406」的下一行不會是單獨數字，但 202411 為 6 位已排除）
        if (prev.Contains('：') && Regex.IsMatch(prev, @"\d{4,6}"))
            return false;

        return true;
    }

    private static bool IsGhostTableRow(string line)
    {
        if (!IsTableRow(line))
            return false;

        var trimmed = line.Trim();
        var cells = trimmed.Split('|', StringSplitOptions.RemoveEmptyEntries)
            .Select(c => c.Trim())
            .ToList();

        if (cells.Count < 2)
            return false;

        if (cells.All(c => Regex.IsMatch(c, @"^[-:\s]+$")))
            return false;

        var nonEmpty = cells.Where(c => !string.IsNullOrWhiteSpace(c)).ToList();
        if (nonEmpty.Count != 1)
            return false;

        return StandalonePageNumber.IsMatch(nonEmpty[0]);
    }

    private static bool IsTableRow(string line) =>
        line.TrimStart().StartsWith('|') && line.TrimEnd().EndsWith('|');

    private static IEnumerable<KnownGap> DetectBrokenTables(string mdContent)
    {
        var lines = mdContent.Split('\n');
        var gaps = new List<KnownGap>();
        var i = 0;

        while (i < lines.Length)
        {
            if (!IsTableRow(lines[i]))
            {
                i++;
                continue;
            }

            var start = i;
            var tableLines = new List<string>();
            while (i < lines.Length && IsTableRow(lines[i]))
            {
                tableLines.Add(lines[i]);
                i++;
            }

            if (tableLines.Count < 2)
                continue;

            var (emptyRatio, totalCells) = ComputeTableEmptyRatio(tableLines);
            if (totalCells == 0)
                continue;

            if (emptyRatio >= 0.55)
            {
                gaps.Add(new KnownGap
                {
                    Type = "broken_table",
                    Location = FindNearestSection(mdContent, mdContent.IndexOf(lines[start], StringComparison.Ordinal)),
                    Description = $"表格空儲存格比例 {emptyRatio:P0}（{tableLines.Count} 列）",
                    Severity = emptyRatio >= 0.7 ? "high" : "moderate"
                });
            }
        }

        return gaps;
    }

    private static string FindNearestSection(string content, int position)
    {
        if (position < 0)
            return "unknown";

        var before = content[..position];
        var lines = before.Split('\n');
        for (var i = lines.Length - 1; i >= 0; i--)
        {
            var trimmed = lines[i].Trim();
            if (string.IsNullOrEmpty(trimmed))
                continue;

            var sectionRef = SectionRefPattern.Match(trimmed);
            if (sectionRef.Success)
            {
                var section = sectionRef.Value.Trim();
                return section.Length > 60 ? section[..60] + "…" : section;
            }

            if (MarkdownHeadingSection.IsMatch(trimmed))
            {
                var heading = trimmed.TrimStart('#').Trim();
                return heading.Length > 60 ? heading[..60] + "…" : heading;
            }

            var bracketRef = BracketSectionRef.Match(trimmed);
            if (bracketRef.Success)
            {
                var section = trimmed.Length > 60 ? trimmed[..60] + "…" : trimmed;
                return section;
            }

            if (SectionHeader.IsMatch(trimmed))
                return trimmed.Length > 60 ? trimmed[..60] + "…" : trimmed;
        }

        return "document";
    }

    private static List<KnownGap> DeduplicateGaps(List<KnownGap> gaps)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var result = new List<KnownGap>();
        foreach (var gap in gaps)
        {
            var key = $"{gap.Type}|{gap.Location}|{gap.Field}";
            if (seen.Add(key))
                result.Add(gap);
        }
        return result;
    }

    private static bool IsDataSourceMarkerLine(string line, out Match markerMatch, out bool isNumbered)
    {
        markerMatch = NumberedDataSourceLine.Match(line);
        if (markerMatch.Success)
        {
            isNumbered = true;
            return true;
        }

        markerMatch = UnnumberedDataSourceLine.Match(line);
        if (markerMatch.Success)
        {
            isNumbered = false;
            return true;
        }

        isNumbered = false;
        markerMatch = Match.Empty;
        return false;
    }

    private static string DataSourceFieldLabel(bool isNumbered) =>
        isNumbered ? "(3) 資料來源" : "資料來源";

    private static string DataSourcePendingTag(bool isNumbered) =>
        isNumbered ? "[待補：(3) 資料來源]" : "[待補：資料來源]";

    private static bool IsDataSourceBoundary(string trimmed)
    {
        if (StructuredFieldMarker.IsMatch(trimmed))
            return true;

        if (RequirementHeader.IsMatch(trimmed.TrimEnd()) ||
            UnnumberedRequirementHeader.IsMatch(trimmed) ||
            FunctionalScreenHeader.IsMatch(trimmed))
            return true;

        if (SectionHeader.IsMatch(trimmed) ||
            SectionRefPattern.IsMatch(trimmed) ||
            BracketSectionRef.IsMatch(trimmed) ||
            MarkdownHeadingSection.IsMatch(trimmed))
            return true;

        return false;
    }

    private static bool IsTrulyBlankDataSource(string[] lines, int markerLineIndex)
    {
        var line = lines[markerLineIndex];
        if (IsDataSourceMarkerLine(line, out var markerMatch, out _))
        {
            var sameLineContent = line[(markerMatch.Index + markerMatch.Length)..].Trim();
            if (IsSubstantiveDataSourceContent(sameLineContent))
                return false;
        }

        for (var i = markerLineIndex + 1; i < lines.Length; i++)
        {
            var trimmed = lines[i].Trim();
            if (string.IsNullOrEmpty(trimmed))
                continue;

            if (IsDataSourceBoundary(trimmed))
                return true;

            if (IsSubstantiveDataSourceContent(trimmed))
                return false;
        }

        return true;
    }

    private static bool IsSubstantiveDataSourceContent(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;

        if (text.StartsWith("> **[待補：(3) 資料來源]", StringComparison.Ordinal) ||
            text.StartsWith("> **[待補：資料來源]", StringComparison.Ordinal))
            return false;

        return true;
    }

    private static bool IsOrphanPageNumberAfterReferenceBlock(string[] lines, int index)
    {
        var trimmed = lines[index].Trim();
        if (!StandalonePageNumber.IsMatch(trimmed))
            return false;

        for (var i = index - 1; i >= 0 && index - i <= 25; i--)
        {
            var prev = lines[i].Trim();
            if (prev.StartsWith("參考資料：", StringComparison.Ordinal) ||
                prev.Contains("[圖片補足：p.", StringComparison.Ordinal))
                return true;

            if (SectionRefPattern.IsMatch(prev) || SectionHeader.IsMatch(prev))
                break;
        }

        return false;
    }

    private static bool IsTrulyBlankRequirementA(string[] lines, int aLineIndex)
    {
        var nextIdx = FindNextNonEmptyLineIndex(lines, aLineIndex + 1);
        if (nextIdx is null)
            return true;

        var next = lines[nextIdx.Value].Trim();
        if (StructuredFieldMarker.IsMatch(next) || SectionHeader.IsMatch(next))
            return true;

        return Regex.IsMatch(next, @"^[B-Z]\.\s*$");
    }

    private static int? FindNextNonEmptyLineIndex(string[] lines, int startIndex)
    {
        for (var i = startIndex; i < lines.Length; i++)
        {
            if (!string.IsNullOrWhiteSpace(lines[i]))
                return i;
        }
        return null;
    }

    private static int LineStartIndex(string content, int lineIndex)
    {
        var lines = content.Split('\n');
        var pos = 0;
        for (var i = 0; i < lineIndex && i < lines.Length; i++)
            pos += lines[i].Length + 1;
        return pos;
    }

    private static string FindReferenceImageLocation(string content, int refIndex)
    {
        if (refIndex < 0)
            return "document";

        var start = Math.Max(0, refIndex - 1200);
        var vicinity = content[start..refIndex];
        var vicinityLines = vicinity.Split('\n');
        for (var i = vicinityLines.Length - 1; i >= 0; i--)
        {
            var trimmed = vicinityLines[i].Trim();
            if (trimmed.Contains("1.1.5", StringComparison.Ordinal) || SectionHeader.IsMatch(trimmed))
                return trimmed.Length > 60 ? trimmed[..60] + "…" : trimmed;
        }

        return FindNearestSection(content, refIndex);
    }

    private static string InsertBlankDataSourceAnnotations(
        string content,
        IReadOnlyList<KnownGap> gaps,
        List<string> log)
    {
        if (!gaps.Any(g => g.Field is "(3) 資料來源" or "資料來源"))
            return content;

        var normalized = content.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n');
        var insertAfter = new Dictionary<int, bool>();

        for (var i = 0; i < lines.Length; i++)
        {
            if (!IsDataSourceMarkerLine(lines[i], out _, out var isNumbered))
                continue;
            if (!IsTrulyBlankDataSource(lines, i))
                continue;

            var pendingTag = DataSourcePendingTag(isNumbered);
            if (i + 1 < lines.Length && lines[i + 1].Contains(pendingTag, StringComparison.Ordinal))
                continue;
            insertAfter[i] = isNumbered;
        }

        if (insertAfter.Count == 0)
            return content;

        var result = new List<string>(lines.Length + insertAfter.Count * 4);
        for (var i = 0; i < lines.Length; i++)
        {
            result.Add(lines[i]);
            if (!insertAfter.TryGetValue(i, out var isNumbered))
                continue;

            var pendingTag = DataSourcePendingTag(isNumbered);
            result.Add("");
            result.Add($"> **{pendingTag}**");
            result.Add("> - 狀態：此段落為空白，請對照 source_file 原檔補充");
            result.Add("> - 置信度：待人工補充");
            log.Add($"插入 {pendingTag} 標記");
        }

        return string.Join(Environment.NewLine, result);
    }

    private static string ApplyBrokenTableAnnotations(
        string content,
        PdfEnhancementResult? pdfEnhancement,
        List<string> log)
    {
        var lines = content.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        var usedTables = new HashSet<int>();
        var insertBefore = new Dictionary<int, List<string>>();

        var i = 0;
        while (i < lines.Length)
        {
            if (!IsTableRow(lines[i]))
            {
                i++;
                continue;
            }

            var start = i;
            var tableLines = new List<string>();
            while (i < lines.Length && IsTableRow(lines[i]))
            {
                tableLines.Add(lines[i]);
                i++;
            }

            if (tableLines.Count < 2)
                continue;

            var (emptyRatio, _) = ComputeTableEmptyRatio(tableLines);
            if (emptyRatio < 0.55)
                continue;

            var preferredPage = InferPageNumberNearLine(lines, start);
            var altTable = FindAlternatePdfTable(pdfEnhancement, usedTables, preferredPage);
            var annotation = BuildBrokenTableAnnotation(emptyRatio, altTable);
            insertBefore[start] = annotation;
            log.Add($"插入 [表格待修] 標記（空儲存格 {emptyRatio:P0}）");
        }

        if (insertBefore.Count == 0)
            return content;

        var result = new List<string>(lines.Length + insertBefore.Values.Sum(v => v.Count));
        for (var j = 0; j < lines.Length; j++)
        {
            if (insertBefore.TryGetValue(j, out var block))
                result.AddRange(block);
            result.Add(lines[j]);
        }

        return string.Join(Environment.NewLine, result);
    }

    private static int? InferPageNumberNearLine(string[] lines, int lineIndex)
    {
        for (var i = lineIndex; i >= 0 && lineIndex - i <= 40; i--)
        {
            var match = PageFractionMarker.Match(lines[i]);
            if (match.Success && int.TryParse(match.Groups[1].Value, out var pageNum))
                return pageNum;
        }

        return null;
    }

    private static PdfTableInfo? FindAlternatePdfTable(
        PdfEnhancementResult? pdfEnhancement,
        HashSet<int> usedTables,
        int? preferredPage)
    {
        if (pdfEnhancement is null || pdfEnhancement.Tables.Count == 0)
            return null;

        if (preferredPage is int page)
        {
            for (var t = 0; t < pdfEnhancement.Tables.Count; t++)
            {
                if (usedTables.Contains(t))
                    continue;
                var candidate = pdfEnhancement.Tables[t];
                if (candidate.IsJunk || candidate.PageNumber != page)
                    continue;
                usedTables.Add(t);
                return candidate;
            }
        }

        for (var t = 0; t < pdfEnhancement.Tables.Count; t++)
        {
            if (usedTables.Contains(t))
                continue;
            if (pdfEnhancement.Tables[t].IsJunk)
                continue;
            usedTables.Add(t);
            return pdfEnhancement.Tables[t];
        }

        return null;
    }

    private static List<string> BuildBrokenTableAnnotation(double emptyRatio, PdfTableInfo? altTable)
    {
        var block = new List<string>
        {
            "",
            "> **[表格待修]**",
            $"> - 狀態：原 Markdown 表格空儲存格比例 {emptyRatio:P0}，請對照 source_file 確認"
        };

        if (altTable is not null)
        {
            block.Add($"> - pdfplumber 擷取備選表格（p.{altTable.PageNumber}）：");
            block.Add(">");
            foreach (var mdLine in altTable.Markdown.Split('\n'))
                block.Add("> " + mdLine);
        }

        block.Add("");
        return block;
    }

    private static string InsertMissingTableFallbacks(
        string content,
        PdfEnhancementResult? pdfEnhancement,
        IReadOnlyList<KnownGap> gaps,
        List<string> log)
    {
        if (pdfEnhancement is null || pdfEnhancement.Pages.Count == 0)
            return content;

        var missingPages = gaps
            .Where(g => g.Type == "missing_table")
            .SelectMany(g =>
            {
                if (g.PdfPages is { Length: > 0 })
                    return g.PdfPages;
                if (g.Location.StartsWith("p.", StringComparison.OrdinalIgnoreCase) &&
                    int.TryParse(g.Location[2..], out var fromLoc))
                    return new[] { fromLoc };
                return Array.Empty<int>();
            })
            .Distinct()
            .OrderBy(p => p)
            .ToList();

        if (missingPages.Count == 0)
            return content;

        var pageLookup = pdfEnhancement.Pages.ToDictionary(p => p.PageNumber);
        var sections = BuildPageSections(content);
        var normalized = content.Replace("\r\n", "\n").Replace('\r', '\n');
        var lines = normalized.Split('\n').ToList();

        // Insert from bottom to top so line indices stay valid
        foreach (var pageNum in missingPages.OrderByDescending(p => p))
        {
            if (!pageLookup.TryGetValue(pageNum, out var page))
                continue;

            if (content.Contains($"[表格遺失：影像表 p.{pageNum}]", StringComparison.Ordinal))
                continue;

            var caption = BuildContextCaptionFromMd(normalized, sections, pageNum, page);
            var block = BuildMissingTableBlock(page, caption);

            int insertAt;
            if (sections.TryGetValue(pageNum, out var span))
            {
                insertAt = Math.Min(span.End, lines.Count);
                // Prefer just before next page marker; skip trailing blanks within section
                while (insertAt > span.Start && insertAt > 0 && string.IsNullOrWhiteSpace(lines[insertAt - 1]))
                    insertAt--;
            }
            else
            {
                insertAt = lines.Count;
            }

            lines.InsertRange(insertAt, block);
            log.Add($"於 p.{pageNum} 插入 [表格遺失] + 圖片描述與截圖");
            if (!string.IsNullOrEmpty(page.OcrStatus) && page.OcrStatus != "not_attempted")
                log.Add($"p.{pageNum} OCR 狀態：{page.OcrStatus}");
        }

        return string.Join(Environment.NewLine, lines);
    }

    private static List<string> BuildMissingTableBlock(PdfPageInfo page, string caption)
    {
        var label = Path.GetFileNameWithoutExtension(page.Filename);
        var block = new List<string>
        {
            "",
            $"> **[表格遺失：影像表 p.{page.PageNumber}]**",
            "> - 狀態：文字層無明細表；已附整頁截圖與文字描述",
            $"> - 圖片描述：{caption}"
        };

        AppendOcrSummaryLines(block, page);

        block.Add($"> {page.Marker}");
        block.Add($"> ![{label}]({page.ImagePath})");
        block.Add("");
        return block;
    }

    private static void AppendOcrSummaryLines(List<string> block, PdfPageInfo page)
    {
        var summary = FormatOcrSummaryLine(page);
        if (string.IsNullOrEmpty(summary))
            return;

        var lines = summary.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        block.Add($"> - OCR 摘要：{lines[0]}");
        for (var i = 1; i < lines.Length; i++)
            block.Add($"> {lines[i]}");
    }

    private static string FormatOcrSummaryLine(PdfPageInfo page)
    {
        if (page.OcrStatus is "ok" or "ok_empty")
        {
            if (string.IsNullOrWhiteSpace(page.OcrPreview))
                return "（OCR 無可用文字）";
            return page.OcrPreview.Replace("\r\n", "\n").Replace('\r', '\n');
        }

        if (page.OcrStatus is "skipped_no_engine" or "skipped_no_executable")
            return "（找不到 Tesseract 執行檔，已略過）";
        if (page.OcrStatus == "skipped_no_language")
            return "（Tesseract 缺少 chi_tra／eng 語言資料，已略過）";
        if (page.OcrStatus == "skipped_no_image")
            return "（無頁面圖片，已略過）";
        if (page.OcrStatus.StartsWith("error", StringComparison.Ordinal))
            return $"（OCR 失敗：{page.OcrStatus}）";
        return "";
    }

    private static string InsertPdfPageImages(
        string content,
        PdfEnhancementResult? pdfEnhancement,
        List<string> log)
    {
        if (pdfEnhancement is null || pdfEnhancement.Pages.Count == 0)
            return content;

        var normalized = content.Replace("\r\n", "\n").Replace('\r', '\n');
        var pageLookup = pdfEnhancement.Pages.ToDictionary(p => p.PageNumber);

        var enhancedCount = 0;
        normalized = ReferenceImagePlaceholder.Replace(normalized, match =>
        {
            if (!int.TryParse(match.Groups[1].Value, out var pageNum) || !pageLookup.TryGetValue(pageNum, out var page))
                return match.Value;

            enhancedCount++;
            return BuildReferenceImageSection(page);
        });

        if (enhancedCount > 0)
            log.Add($"於 {enhancedCount} 處「參考資料」插入 【imgN】 標記與圖片連結");

        var indexBlock = BuildImageIndexSection(pdfEnhancement.Pages);
        log.Add($"於文件末尾附加圖片索引（共 {pdfEnhancement.Pages.Count} 張；描述與索引僅寫入最終 .md，不寫入 .raw.md）");
        return normalized.TrimEnd() + Environment.NewLine + indexBlock;
    }

    private static string BuildReferenceImageSection(PdfPageInfo page)
    {
        var label = Path.GetFileNameWithoutExtension(page.Filename);
        return $"""
            參考資料：
            {page.Marker}
            > **[圖片補足：p.{page.PageNumber}]**
            > {page.Marker}
            > ![{label}]({page.ImagePath})
            """;
    }

    private static string BuildImageBlock(PdfPageInfo page, bool includeStatus = true)
    {
        var label = Path.GetFileNameWithoutExtension(page.Filename);
        var sb = new StringBuilder();
        sb.AppendLine($"> **[圖片補足：p.{page.PageNumber}]**");
        if (!string.IsNullOrWhiteSpace(page.ImageCaption))
            sb.AppendLine($"> - 圖片描述：{page.ImageCaption}");
        sb.AppendLine($"> {page.Marker}");
        sb.AppendLine($"> ![{label}]({page.ImagePath})");
        if (includeStatus)
            sb.AppendLine("> - 狀態：已自動擷取 PDF 頁面圖片，請人工補充欄位說明");
        sb.AppendLine();
        return sb.ToString();
    }

    private static string BuildImageIndexSection(IReadOnlyList<PdfPageInfo> pages)
    {
        var sb = new StringBuilder();
        sb.AppendLine();
        sb.AppendLine("## 圖片索引");
        sb.AppendLine();
        foreach (var page in pages)
            sb.Append(BuildImageBlock(page));
        return sb.ToString();
    }

    private static (double emptyRatio, int totalCells) ComputeTableEmptyRatio(IReadOnlyList<string> tableLines)
    {
        var totalCells = 0;
        var emptyCells = 0;
        foreach (var row in tableLines)
        {
            if (row.Contains("---"))
                continue;

            var cells = row.Split('|', StringSplitOptions.RemoveEmptyEntries);
            foreach (var cell in cells)
            {
                totalCells++;
                if (string.IsNullOrWhiteSpace(cell))
                    emptyCells++;
            }
        }

        if (totalCells == 0)
            return (0, 0);

        return ((double)emptyCells / totalCells, totalCells);
    }

    private static string InferDocumentType(string sourcePath) =>
        Path.GetExtension(sourcePath).ToLowerInvariant() switch
        {
            ".pdf" => "pdf",
            ".doc" or ".docx" => "word",
            ".xls" or ".xlsx" => "spreadsheet",
            ".ppt" or ".pptx" => "presentation",
            _ => "document"
        };

    private static string StripExistingFrontMatter(string content)
    {
        if (!content.StartsWith("---", StringComparison.Ordinal))
            return content;

        var end = content.IndexOf("\n---", 3, StringComparison.Ordinal);
        if (end < 0)
            return content;

        return content[(end + 4)..];
    }

    private static string YamlQuote(string value) =>
        "\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";

    private static string TruncateForLog(string line) =>
        line.Trim().Length > 80 ? line.Trim()[..80] + "…" : line.Trim();

    private static void WriteQualityJson(
        string path,
        MarkdownMetadata metadata,
        IReadOnlyList<string> cleanupLog,
        IReadOnlyList<string> enhancementLog,
        IReadOnlyList<KnownGap> gaps,
        PdfEnhancementResult? pdfEnhancement,
        DocxConversionResult? docxConversion = null,
        XlsxConversionResult? xlsxConversion = null)
    {
        var report = new
        {
            source_file = metadata.SourceFile,
            conversion_engine = metadata.ConversionEngine,
            converted_date = metadata.ConvertedDate.ToString("yyyy-MM-ddTHH:mm:ss"),
            cleanup = new
            {
                actions_count = cleanupLog.Count,
                actions = cleanupLog
            },
            openxml_images = docxConversion is null || docxConversion.Images.Count == 0 ? null : new
            {
                count = docxConversion.ImageCount,
                images_folder = docxConversion.ImagesFolder,
                image_paths = docxConversion.ImagePaths,
                export_log = docxConversion.ExportLog,
                images = docxConversion.Images.Select(img => new
                {
                    index = img.Index,
                    filename = img.Filename,
                    marker = img.Marker,
                    image_path = img.ImagePath
                })
            },
            openxml_sheets = xlsxConversion is null || xlsxConversion.SheetCount == 0 ? null : new
            {
                count = xlsxConversion.SheetCount,
                sheet_names = xlsxConversion.SheetNames,
                conversion_log = xlsxConversion.ConversionLog
            },
            pdf_enhancement = pdfEnhancement is null ? null : new
            {
                pages_exported = pdfEnhancement.PagesExported > 0
                    ? pdfEnhancement.PagesExported
                    : pdfEnhancement.Pages.Count,
                tables_extracted = pdfEnhancement.TablesExtractedTotal > 0
                    ? pdfEnhancement.TablesExtractedTotal
                    : (pdfEnhancement.TablesExtracted > 0
                        ? pdfEnhancement.TablesExtracted
                        : pdfEnhancement.Tables.Count),
                tables_extracted_total = pdfEnhancement.TablesExtractedTotal > 0
                    ? pdfEnhancement.TablesExtractedTotal
                    : pdfEnhancement.Tables.Count,
                tables_extracted_good = pdfEnhancement.TablesExtractedGood,
                images_folder = pdfEnhancement.ImagesFolder,
                image_paths = pdfEnhancement.ImagePaths.Count > 0
                    ? pdfEnhancement.ImagePaths
                    : pdfEnhancement.Pages.Select(p => p.ImagePath).Where(p => !string.IsNullOrEmpty(p)).ToList(),
                pages = pdfEnhancement.Pages.Select(p => new
                {
                    index = p.Index,
                    filename = p.Filename,
                    page = p.PageNumber,
                    marker = p.Marker,
                    image_path = p.ImagePath,
                    char_count = p.CharCount,
                    embedded_image_count = p.EmbeddedImageCount,
                    content_image_count = p.ContentImageCount,
                    good_table_count = p.GoodTableCount,
                    junk_table_count = p.JunkTableCount,
                    likely_image_table = p.LikelyImageTable,
                    image_caption = p.ImageCaption,
                    ocr_status = p.OcrStatus,
                    ocr_preview_length = string.IsNullOrEmpty(p.OcrPreview) ? 0 : p.OcrPreview.Length
                }),
                note = "圖片索引與缺表文字描述僅寫入最終 .md，不會出現在 .raw.md",
                errors = pdfEnhancement.Errors,
                error = pdfEnhancement.Error
            },
            enhancements = new
            {
                actions_count = enhancementLog.Count,
                actions = enhancementLog
            },
            gaps = gaps.Select(g => new
            {
                g.Type,
                g.Location,
                g.Field,
                g.Description,
                g.Severity,
                pdf_pages = g.PdfPages
            })
        };

        var json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json, Encoding.UTF8);
    }
}
