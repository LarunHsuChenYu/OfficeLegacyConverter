namespace OfficeLegacyConverter;

internal sealed class MarkdownMetadata
{
    public string SourceFile { get; init; } = "";
    public string SourcePath { get; init; } = "";
    public string ConvertedFile { get; init; } = "";
    public DateTime ConvertedDate { get; init; } = DateTime.Now;
    public string Converter { get; init; } = "OfficeLegacyConverter + markitdown";
    public string ConversionEngine { get; init; } = "MarkItDown";
    public string? ConverterVersion { get; init; }
    public int KnowledgeTier { get; init; } = 1;
    public string DocumentType { get; init; } = "document";
    public string? ImagesFolder { get; init; }
    public IReadOnlyList<KnownGap> KnownGaps { get; init; } = [];
}

internal sealed class KnownGap
{
    public string Type { get; init; } = "";
    public string Location { get; init; } = "";
    public string? Field { get; init; }
    public string? Description { get; init; }
    public string? Severity { get; init; }
    public int[]? PdfPages { get; init; }
}

internal sealed class PostProcessResult
{
    public string FinalMdPath { get; init; } = "";
    public string? RawMdPath { get; init; }
    public string? QualityJsonPath { get; init; }
    public IReadOnlyList<string> CleanupLog { get; init; } = [];
    public IReadOnlyList<KnownGap> DetectedGaps { get; init; } = [];
    public PdfEnhancementResult? PdfEnhancement { get; init; }
}
