using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using Wp = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeLegacyConverter;

internal sealed class DocxConversionResult
{
    public string Markdown { get; init; } = "";
    public string? ImagesFolder { get; init; }
    public IReadOnlyList<DocxImageInfo> Images { get; init; } = [];
    public IReadOnlyList<string> ImagePaths => Images.Select(i => i.ImagePath).ToList();
    public int ImageCount => Images.Count;
    public IReadOnlyList<string> ExportLog { get; init; } = [];
}

internal sealed class DocxImageInfo
{
    public int Index { get; init; }
    public string Filename { get; init; } = "";
    public string Marker { get; init; } = "";
    public string ImagePath { get; init; } = "";
}

internal static class DocxMarkdownConverter
{
    private const string VmlNamespace = "urn:schemas-microsoft-com:vml";
    private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static DocxConversionResult Convert(
        string docxPath,
        string outputDir,
        string baseName,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        var imagesFolderPath = Path.Combine(outputDir, baseName);
        var sb = new StringBuilder();
        var exportLog = new List<string>();

        using var document = WordprocessingDocument.Open(docxPath, false);
        var mainPart = document.MainDocumentPart
            ?? throw new InvalidOperationException("無法讀取 Word 文件內容。");
        var body = mainPart.Document?.Body
            ?? throw new InvalidOperationException("無法讀取 Word 文件內容。");

        var imageExporter = new DocxImageExporter(mainPart, imagesFolderPath, baseName);

        foreach (var element in body.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();

            switch (element)
            {
                case Paragraph paragraph:
                    AppendParagraph(sb, paragraph, imageExporter);
                    break;
                case Table table:
                    AppendTable(sb, table, imageExporter);
                    break;
            }
        }

        var inlineCount = imageExporter.InlineExportCount;
        imageExporter.FallbackExportFromMedia(docxPath, exportLog, cancellationToken);
        var images = imageExporter.ExportedImages;

        if (images.Count > 0)
            AppendImageIndex(sb, images, baseName);

        if (images.Count > 0 && exportLog.Count == 0)
            exportLog.Add($"內嵌解析匯出 {inlineCount} 張圖片（DrawingML/VML）");

        return new DocxConversionResult
        {
            Markdown = sb.ToString().TrimEnd() + Environment.NewLine,
            ImagesFolder = images.Count > 0 ? $"{baseName}/" : null,
            Images = images,
            ExportLog = exportLog
        };
    }

    private static void AppendParagraph(StringBuilder sb, Paragraph paragraph, DocxImageExporter imageExporter)
    {
        var textParts = new List<string>();
        foreach (var child in paragraph.Elements())
        {
            if (child is Run run)
            {
                foreach (var runChild in run.Elements())
                {
                    switch (runChild)
                    {
                        case Text textNode:
                            textParts.Add(textNode.Text);
                            break;
                        case Break:
                            textParts.Add(Environment.NewLine);
                            break;
                        case Wp.Inline inline:
                            AppendInlineImage(textParts, inline, imageExporter);
                            break;
                        case Wp.Anchor anchor:
                            AppendAnchorImage(textParts, anchor, imageExporter);
                            break;
                        case Picture pict:
                            AppendVmlImage(textParts, pict, imageExporter);
                            break;
                        default:
                            if (IsPictElement(runChild))
                                AppendVmlImage(textParts, runChild, imageExporter);
                            break;
                    }
                }
            }
            else if (child is Wp.Inline inline)
            {
                AppendInlineImage(textParts, inline, imageExporter);
            }
            else if (child is Wp.Anchor anchor)
            {
                AppendAnchorImage(textParts, anchor, imageExporter);
            }
            else if (child is Picture pict)
            {
                AppendVmlImage(textParts, pict, imageExporter);
            }
            else if (IsPictElement(child))
            {
                AppendVmlImage(textParts, child, imageExporter);
            }
        }

        var text = string.Concat(textParts).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            sb.AppendLine();
            return;
        }

        var headingLevel = GetHeadingLevel(paragraph);
        if (headingLevel > 0)
        {
            sb.Append(new string('#', Math.Clamp(headingLevel, 1, 6)));
            sb.Append(' ');
            sb.AppendLine(EscapeMarkdownInline(text));
            return;
        }

        sb.AppendLine(EscapeMarkdownInline(text));
    }

    private static void AppendInlineImage(List<string> parts, Wp.Inline inline, DocxImageExporter exporter)
    {
        var info = exporter.ExportFromDrawing(inline);
        if (info is not null)
            parts.Add(BuildImageMarkdown(info));
    }

    private static void AppendAnchorImage(List<string> parts, Wp.Anchor anchor, DocxImageExporter exporter)
    {
        var info = exporter.ExportFromDrawing(anchor);
        if (info is not null)
            parts.Add(BuildImageMarkdown(info));
    }

    private static void AppendVmlImage(List<string> parts, OpenXmlElement container, DocxImageExporter exporter)
    {
        foreach (var info in exporter.ExportFromVml(container))
            parts.Add(BuildImageMarkdown(info));
    }

    private static string BuildImageMarkdown(DocxImageInfo info) =>
        $"{info.Marker}\n\n![{Path.GetFileNameWithoutExtension(info.Filename)}]({info.ImagePath})\n";

    private static void AppendTable(StringBuilder sb, Table table, DocxImageExporter imageExporter)
    {
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count == 0)
            return;

        var grid = BuildTableGrid(rows, imageExporter);
        if (grid.Count == 0)
            return;

        var columnCount = grid.Max(r => r.Count);
        if (columnCount == 0)
            return;

        sb.AppendLine();
        for (var r = 0; r < grid.Count; r++)
        {
            var row = grid[r];
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

    private static List<List<string>> BuildTableGrid(IReadOnlyList<TableRow> rows, DocxImageExporter imageExporter)
    {
        var grid = new List<List<string>>();
        var verticalCarry = new Dictionary<int, string>();

        foreach (var row in rows)
        {
            var cells = new List<string>();
            var column = 0;

            foreach (var cell in row.Elements<TableCell>())
            {
                while (verticalCarry.ContainsKey(column))
                {
                    cells.Add("");
                    column++;
                }

                var gridSpan = (int)(cell.TableCellProperties?.GridSpan?.Val?.Value ?? 1);
                var vMerge = cell.TableCellProperties?.VerticalMerge?.Val?.Value;
                var text = GetCellText(cell, imageExporter);

                if (vMerge == MergedCellValues.Continue)
                {
                    cells.Add("");
                    for (var span = 1; span < gridSpan; span++)
                    {
                        cells.Add("");
                        column++;
                    }
                }
                else
                {
                    cells.Add(text);
                    verticalCarry[column] = text;
                    for (var span = 1; span < gridSpan; span++)
                    {
                        cells.Add("");
                        column++;
                    }
                }

                column += gridSpan;
            }

            while (verticalCarry.Keys.Any(k => k >= cells.Count))
                cells.Add("");

            grid.Add(cells);
        }

        return grid;
    }

    private static string GetCellText(TableCell cell, DocxImageExporter imageExporter)
    {
        var parts = new List<string>();
        foreach (var paragraph in cell.Descendants<Paragraph>())
        {
            foreach (var child in paragraph.Elements())
            {
                if (child is Run run)
                {
                    foreach (var runChild in run.Elements())
                    {
                        switch (runChild)
                        {
                            case Text textNode:
                                parts.Add(textNode.Text);
                                break;
                            case Wp.Inline inline:
                                AppendInlineImage(parts, inline, imageExporter);
                                break;
                            case Wp.Anchor anchor:
                                AppendAnchorImage(parts, anchor, imageExporter);
                                break;
                            case Picture pict:
                                AppendVmlImage(parts, pict, imageExporter);
                                break;
                            default:
                                if (IsPictElement(runChild))
                                    AppendVmlImage(parts, runChild, imageExporter);
                                break;
                        }
                    }
                }
                else if (child is Wp.Inline inline)
                {
                    AppendInlineImage(parts, inline, imageExporter);
                }
                else if (child is Wp.Anchor anchor)
                {
                    AppendAnchorImage(parts, anchor, imageExporter);
                }
                else if (child is Picture pict)
                {
                    AppendVmlImage(parts, pict, imageExporter);
                }
                else if (IsPictElement(child))
                {
                    AppendVmlImage(parts, child, imageExporter);
                }
            }
        }

        return string.Join(" ", parts.Select(p => p.Replace('\n', ' '))).Trim();
    }

    private static int GetHeadingLevel(Paragraph paragraph)
    {
        var props = paragraph.ParagraphProperties;
        if (props is null)
            return 0;

        var styleId = props.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrEmpty(styleId))
        {
            if (styleId.StartsWith("Heading", StringComparison.OrdinalIgnoreCase)
                && int.TryParse(styleId.AsSpan(7), out var headingLevel))
                return headingLevel;

            if (styleId.StartsWith("标题", StringComparison.OrdinalIgnoreCase)
                || styleId.StartsWith("標題", StringComparison.OrdinalIgnoreCase))
            {
                var digits = new string(styleId.Where(char.IsDigit).ToArray());
                if (int.TryParse(digits, out var cnLevel))
                    return cnLevel;
            }
        }

        var outline = props.GetFirstChild<OutlineLevel>()?.Val?.Value;
        if (outline.HasValue)
            return Math.Clamp(outline.Value + 1, 1, 6);

        return 0;
    }

    private static string EscapeMarkdownInline(string text) =>
        text.Replace("\r", "").Replace("\n", " ").Trim();

    private static string EscapeTableCell(string text) =>
        text.Replace("|", "\\|", StringComparison.Ordinal).Replace("\n", " ", StringComparison.Ordinal).Trim();

    private static void AppendImageIndex(StringBuilder sb, IReadOnlyList<DocxImageInfo> images, string baseName)
    {
        sb.AppendLine();
        sb.AppendLine("## 圖片索引");
        sb.AppendLine();
        foreach (var image in images)
        {
            sb.AppendLine($"> {image.Marker}");
            sb.AppendLine($"> ![{Path.GetFileNameWithoutExtension(image.Filename)}]({image.ImagePath})");
            sb.AppendLine();
        }
    }

    private static bool IsPictElement(OpenXmlElement element) =>
        string.Equals(element.LocalName, "pict", StringComparison.Ordinal)
        && string.Equals(element.NamespaceUri, "http://schemas.openxmlformats.org/wordprocessingml/2006/main", StringComparison.Ordinal);

    private static IEnumerable<string> FindVmlRelationshipIds(OpenXmlElement element)
    {
        foreach (var desc in element.Descendants())
        {
            if (!string.Equals(desc.LocalName, "imagedata", StringComparison.Ordinal))
                continue;

            if (!string.Equals(desc.NamespaceUri, VmlNamespace, StringComparison.Ordinal)
                && !string.Equals(desc.Prefix, "v", StringComparison.Ordinal))
                continue;

            foreach (var attr in desc.GetAttributes())
            {
                if (!string.Equals(attr.LocalName, "id", StringComparison.Ordinal))
                    continue;
                if (!string.Equals(attr.NamespaceUri, RelationshipNamespace, StringComparison.Ordinal)
                    && !string.Equals(attr.Prefix, "r", StringComparison.Ordinal))
                    continue;

                if (!string.IsNullOrEmpty(attr.Value))
                    yield return attr.Value;
            }
        }
    }

    private sealed class DocxImageExporter
    {
        private readonly MainDocumentPart _mainPart;
        private readonly string _imagesFolderPath;
        private readonly string _baseName;
        private readonly Dictionary<string, DocxImageInfo> _exportedByRelId = new(StringComparer.Ordinal);
        private readonly HashSet<string> _exportedPartUris = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<DocxImageInfo> _allImages = [];
        private int _nextIndex = 1;

        public DocxImageExporter(MainDocumentPart mainPart, string imagesFolderPath, string baseName)
        {
            _mainPart = mainPart;
            _imagesFolderPath = imagesFolderPath;
            _baseName = baseName;
        }

        public int InlineExportCount => _exportedByRelId.Count;

        public IReadOnlyList<DocxImageInfo> ExportedImages =>
            _allImages.OrderBy(i => i.Index).ToList();

        public DocxImageInfo? ExportFromDrawing(OpenXmlElement drawingContainer)
        {
            var blip = drawingContainer.Descendants<A.Blip>().FirstOrDefault();
            if (blip?.Embed?.Value is not { } relId)
                return null;

            return ExportByRelationshipId(relId);
        }

        public IEnumerable<DocxImageInfo> ExportFromVml(OpenXmlElement container)
        {
            foreach (var relId in FindVmlRelationshipIds(container))
            {
                var info = ExportByRelationshipId(relId);
                if (info is not null)
                    yield return info;
            }
        }

        public void FallbackExportFromMedia(string docxPath, List<string> log, CancellationToken cancellationToken)
        {
            var inlineCount = InlineExportCount;
            var fallbackCount = 0;

            using (var zip = ZipFile.Open(docxPath, ZipArchiveMode.Read))
            {
                var mediaFiles = zip.Entries
                    .Where(e => e.FullName.StartsWith("word/media/", StringComparison.OrdinalIgnoreCase)
                                && !string.IsNullOrEmpty(e.Name))
                    .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                foreach (var entry in mediaFiles)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    var partUri = NormalizePartUri(entry.FullName);
                    if (_exportedPartUris.Contains(partUri))
                        continue;

                    Directory.CreateDirectory(_imagesFolderPath);
                    var index = _nextIndex++;
                    var extension = GetExtensionFromPath(entry.Name);
                    var filename = $"img{index}{extension}";
                    var absolutePath = Path.Combine(_imagesFolderPath, filename);

                    using (var stream = entry.Open())
                    using (var file = File.Create(absolutePath))
                        stream.CopyTo(file);

                    var info = new DocxImageInfo
                    {
                        Index = index,
                        Filename = filename,
                        Marker = $"【img{index}】",
                        ImagePath = $"{_baseName}/{filename}"
                    };
                    _allImages.Add(info);
                    _exportedPartUris.Add(partUri);
                    fallbackCount++;
                }
            }

            if (inlineCount > 0 && fallbackCount > 0)
                log.Add($"內嵌解析匯出 {inlineCount} 張、後備媒體匯出 {fallbackCount} 張圖片");
            else if (fallbackCount > 0)
                log.Add($"後備媒體匯出 {fallbackCount} 張圖片（未偵測到內嵌參照）");
            else if (inlineCount > 0)
                log.Add($"內嵌解析匯出 {inlineCount} 張圖片（DrawingML/VML）");
        }

        private DocxImageInfo? ExportByRelationshipId(string relId)
        {
            if (_exportedByRelId.TryGetValue(relId, out var existing))
                return existing;

            OpenXmlPart part;
            try
            {
                part = _mainPart.GetPartById(relId);
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }

            if (part is not ImagePart imagePart)
                return null;

            Directory.CreateDirectory(_imagesFolderPath);
            var index = _nextIndex++;
            var extension = GetImageExtension(imagePart);
            var filename = $"img{index}{extension}";
            var absolutePath = Path.Combine(_imagesFolderPath, filename);

            using (var stream = imagePart.GetStream())
            using (var file = File.Create(absolutePath))
                stream.CopyTo(file);

            var info = new DocxImageInfo
            {
                Index = index,
                Filename = filename,
                Marker = $"【img{index}】",
                ImagePath = $"{_baseName}/{filename}"
            };
            _exportedByRelId[relId] = info;
            _allImages.Add(info);
            _exportedPartUris.Add(NormalizePartUri(imagePart.Uri.OriginalString));
            return info;
        }

        private static string NormalizePartUri(string uri) =>
            uri.StartsWith('/') ? uri : "/" + uri.Replace('\\', '/');
    }

    private static string GetImageExtension(ImagePart part)
    {
        var contentType = part.ContentType.ToLowerInvariant();
        return contentType switch
        {
            "image/png" => ".png",
            "image/jpeg" or "image/jpg" => ".jpg",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tiff",
            "image/x-emf" or "image/emf" => ".emf",
            "image/x-wmf" or "image/wmf" => ".wmf",
            _ => Path.GetExtension(part.Uri.OriginalString)
        } is { Length: > 0 } ext ? ext : ".png";
    }

    private static string GetExtensionFromPath(string fileName)
    {
        var ext = Path.GetExtension(fileName);
        if (string.IsNullOrEmpty(ext))
            return ".png";

        return ext.ToLowerInvariant() switch
        {
            ".png" => ".png",
            ".jpeg" or ".jpg" => ".jpg",
            ".gif" => ".gif",
            ".bmp" => ".bmp",
            ".tiff" or ".tif" => ".tiff",
            ".emf" => ".emf",
            ".wmf" => ".wmf",
            _ => ext
        };
    }
}
