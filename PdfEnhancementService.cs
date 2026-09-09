using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeLegacyConverter;

internal sealed class PdfEnhancementResult
{
    public int PagesExported { get; init; }
    public int TablesExtracted { get; init; }
    public int TablesExtractedTotal { get; init; }
    public int TablesExtractedGood { get; init; }
    public string? ImagesFolder { get; init; }
    public IReadOnlyList<PdfPageInfo> Pages { get; init; } = [];
    public IReadOnlyList<PdfTableInfo> Tables { get; init; } = [];
    public IReadOnlyList<string> ImagePaths { get; init; } = [];
    public IReadOnlyList<string> Errors { get; init; } = [];
    public string? Error { get; init; }
}

internal sealed class PdfPageInfo
{
    public int Index { get; init; }
    public string Filename { get; init; } = "";
    public int PageNumber { get; init; }
    public string Marker { get; init; } = "";
    public string ImagePath { get; init; } = "";
    public int CharCount { get; init; }
    public int EmbeddedImageCount { get; init; }
    public int ContentImageCount { get; init; }
    public int GoodTableCount { get; init; }
    public int JunkTableCount { get; init; }
    public bool LikelyImageTable { get; init; }
    public string ImageCaption { get; init; } = "";
    public string OcrPreview { get; init; } = "";
    public string OcrStatus { get; init; } = "not_attempted";
    public string OcrMethod { get; init; } = "none";
    public IReadOnlyList<string> ContextHints { get; init; } = [];
}

internal sealed class PdfTableInfo
{
    public int PageNumber { get; init; }
    public int TableIndex { get; init; }
    public string Markdown { get; init; } = "";
    public double EmptyCellRatio { get; init; }
    public int RowCount { get; init; }
    public bool IsJunk { get; init; }
}

internal static class PdfEnhancementService
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
    };

    public static PdfEnhancementResult Enhance(
        string pdfPath,
        string outputDir,
        string basename,
        CancellationToken cancellationToken = default)
    {
        var scriptPath = ResolveScriptPath();
        if (scriptPath is null)
        {
            return new PdfEnhancementResult
            {
                Error = "找不到 scripts/pdf_enhance.py"
            };
        }

        if (MarkItDownConverter.ResolvedPythonCommand is null)
        {
            return new PdfEnhancementResult
            {
                Error = "Python 環境未就緒，略過 PDF 增強"
            };
        }

        var arguments = $"\"{scriptPath}\" \"{pdfPath}\" \"{outputDir}\" \"{basename}\"";
        try
        {
            var stdout = RunPython(arguments, cancellationToken);
            return ParseResult(stdout);
        }
        catch (Exception ex)
        {
            return new PdfEnhancementResult
            {
                Error = ex.Message
            };
        }
    }

    private static string? ResolveScriptPath()
    {
        var candidates = new List<string>
        {
            Path.Combine(AppContext.BaseDirectory, "scripts", "pdf_enhance.py")
        };

        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 6; i++)
        {
            candidates.Add(Path.Combine(dir, "scripts", "pdf_enhance.py"));
            var parent = Directory.GetParent(dir);
            if (parent is null)
                break;
            dir = parent.FullName;
        }

        return candidates.FirstOrDefault(File.Exists);
    }

    private static string RunPython(string arguments, CancellationToken cancellationToken)
    {
        var command = MarkItDownConverter.ResolvedPythonCommand
            ?? throw new InvalidOperationException(MarkItDownConverter.EnvironmentStatus);

        var (fileName, args) = SplitCommand(command, arguments);

        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = System.Text.Encoding.UTF8,
                StandardErrorEncoding = System.Text.Encoding.UTF8
            }
        };

        process.Start();

        using var registration = cancellationToken.Register(() =>
        {
            try
            {
                if (!process.HasExited)
                    process.Kill(entireProcessTree: true);
            }
            catch { /* ignore */ }
        });

        var stdout = process.StandardOutput.ReadToEnd();
        var stderr = process.StandardError.ReadToEnd();
        process.WaitForExit();

        cancellationToken.ThrowIfCancellationRequested();

        if (process.ExitCode != 0 && string.IsNullOrWhiteSpace(stdout))
            throw new InvalidOperationException(
                $"pdf_enhance.py 結束代碼 {process.ExitCode}" +
                (string.IsNullOrWhiteSpace(stderr) ? "" : $"\n{stderr.Trim()}"));

        return stdout;
    }

    private static (string FileName, string Arguments) SplitCommand(string command, string arguments)
    {
        var parts = command.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
        var fileName = parts[0];
        var commandArgs = parts.Length > 1 ? parts[1] + " " : "";
        return (fileName, commandArgs + arguments);
    }

    private static PdfEnhancementResult ParseResult(string stdout)
    {
        if (string.IsNullOrWhiteSpace(stdout))
        {
            return new PdfEnhancementResult { Error = "pdf_enhance.py 無輸出" };
        }

        try
        {
            var json = JsonSerializer.Deserialize<PdfEnhanceJson>(stdout.Trim(), JsonOptions);
            if (json is null)
                return new PdfEnhancementResult { Error = "無法解析 pdf_enhance.py JSON" };

            var pages = json.Pages?.Select(p =>
                {
                    var pageNumber = p.PageNumber > 0 ? p.PageNumber : p.Page;
                    var index = p.Index > 0 ? p.Index : pageNumber;
                    return new PdfPageInfo
                    {
                        Index = index,
                        Filename = p.Filename ?? $"img{index}.png",
                        PageNumber = pageNumber,
                        Marker = p.Marker ?? $"【img{index}】",
                        ImagePath = p.ImagePath ?? "",
                        CharCount = p.CharCount,
                        EmbeddedImageCount = p.EmbeddedImageCount,
                        ContentImageCount = p.ContentImageCount,
                        GoodTableCount = p.GoodTableCount,
                        JunkTableCount = p.JunkTableCount,
                        LikelyImageTable = p.LikelyImageTable,
                        ImageCaption = p.ImageCaption ?? "",
                        OcrPreview = p.OcrPreview ?? "",
                        OcrStatus = p.OcrStatus ?? "not_attempted",
                        OcrMethod = p.OcrMethod ?? "none",
                        ContextHints = p.ContextHints ?? []
                    };
                }).ToList() ?? [];
            var tables = json.Tables?.Select(t => new PdfTableInfo
            {
                PageNumber = t.PageNumber,
                TableIndex = t.TableIndex,
                Markdown = t.Markdown ?? "",
                EmptyCellRatio = t.EmptyCellRatio,
                RowCount = t.RowCount,
                IsJunk = t.IsJunk
            }).ToList() ?? [];
            var imagePaths = json.ImagePaths is { Count: > 0 }
                ? json.ImagePaths
                : pages.Select(p => p.ImagePath).Where(p => !string.IsNullOrEmpty(p)).ToList();

            var total = json.TablesExtractedTotal > 0
                ? json.TablesExtractedTotal
                : (json.TablesExtracted > 0 ? json.TablesExtracted : tables.Count);
            var good = json.TablesExtractedGood > 0
                ? json.TablesExtractedGood
                : tables.Count(t => !t.IsJunk);

            return new PdfEnhancementResult
            {
                PagesExported = json.PagesExported > 0 ? json.PagesExported : pages.Count,
                TablesExtracted = total,
                TablesExtractedTotal = total,
                TablesExtractedGood = good,
                ImagesFolder = json.ImagesFolder,
                Pages = pages,
                Tables = tables,
                ImagePaths = imagePaths,
                Errors = json.Errors ?? [],
                Error = json.Error
            };
        }
        catch (JsonException ex)
        {
            return new PdfEnhancementResult
            {
                Error = $"JSON 解析失敗：{ex.Message}"
            };
        }
    }

    private sealed class PdfEnhanceJson
    {
        [JsonPropertyName("pages_exported")]
        public int PagesExported { get; set; }

        [JsonPropertyName("tables_extracted")]
        public int TablesExtracted { get; set; }

        [JsonPropertyName("tables_extracted_total")]
        public int TablesExtractedTotal { get; set; }

        [JsonPropertyName("tables_extracted_good")]
        public int TablesExtractedGood { get; set; }

        [JsonPropertyName("images_folder")]
        public string? ImagesFolder { get; set; }

        public List<PdfPageJson>? Pages { get; set; }
        public List<PdfTableJson>? Tables { get; set; }

        [JsonPropertyName("image_paths")]
        public List<string>? ImagePaths { get; set; }
        public List<string>? Errors { get; set; }
        public string? Error { get; set; }
    }

    private sealed class PdfPageJson
    {
        public int Index { get; set; }

        public string? Filename { get; set; }

        public int Page { get; set; }

        [JsonPropertyName("page_number")]
        public int PageNumber { get; set; }

        public string? Marker { get; set; }

        [JsonPropertyName("image_path")]
        public string? ImagePath { get; set; }

        [JsonPropertyName("char_count")]
        public int CharCount { get; set; }

        [JsonPropertyName("embedded_image_count")]
        public int EmbeddedImageCount { get; set; }

        [JsonPropertyName("content_image_count")]
        public int ContentImageCount { get; set; }

        [JsonPropertyName("good_table_count")]
        public int GoodTableCount { get; set; }

        [JsonPropertyName("junk_table_count")]
        public int JunkTableCount { get; set; }

        [JsonPropertyName("likely_image_table")]
        public bool LikelyImageTable { get; set; }

        [JsonPropertyName("image_caption")]
        public string? ImageCaption { get; set; }

        [JsonPropertyName("ocr_preview")]
        public string? OcrPreview { get; set; }

        [JsonPropertyName("ocr_status")]
        public string? OcrStatus { get; set; }

        [JsonPropertyName("ocr_method")]
        public string? OcrMethod { get; set; }

        [JsonPropertyName("context_hints")]
        public List<string>? ContextHints { get; set; }
    }

    private sealed class PdfTableJson
    {
        [JsonPropertyName("page_number")]
        public int PageNumber { get; set; }

        [JsonPropertyName("table_index")]
        public int TableIndex { get; set; }

        public string? Markdown { get; set; }

        [JsonPropertyName("empty_cell_ratio")]
        public double EmptyCellRatio { get; set; }

        [JsonPropertyName("row_count")]
        public int RowCount { get; set; }

        [JsonPropertyName("is_junk")]
        public bool IsJunk { get; set; }
    }
}
