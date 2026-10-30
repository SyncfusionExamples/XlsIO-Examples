#nullable enable

namespace AgentRAGWithAzure.Models;

public sealed class LocalChunkRecord
{
    public string Id { get; set; } =
        string.Empty;

    public string SourceName { get; set; } =
        string.Empty;

    public string FileExtension { get; set; } =
        string.Empty;

    public int ChunkIndex { get; set; }

    public string Content { get; set; } =
        string.Empty;

    public float[] Embedding { get; set; } =
        Array.Empty<float>();

    public IReadOnlyDictionary<string, object>
        MetadataAttributes
    { get; set; } =
            new Dictionary<string, object>();

    public string CitationDisplay { get; set; } =
        string.Empty;

    public IReadOnlyDictionary<string, object>
        CitationDetails
    { get; set; } =
            new Dictionary<string, object>();

    // Retrieval-time only. This is not stored in PostgreSQL.
    public double Score { get; set; }
}

public sealed class UploadResponse
{
    public List<string> Files { get; set; } =
        new();
}

public sealed class UploadErrorResponse
{
    public string Error { get; set; } =
        string.Empty;
}

public sealed class ChatRequest
{
    public string Question { get; set; } =
        string.Empty;
}

public sealed class ChatResponse
{
    public string Answer { get; set; } =
        string.Empty;

    public List<CitationDto> Citations { get; set; } =
        new();
}

public sealed class ChatErrorResponse
{
    public string Error { get; set; } =
        string.Empty;
}

public sealed class ChatAnswerResult
{
    public string Answer { get; set; } =
        string.Empty;

    public List<CitationDto> Citations { get; set; } =
        new();
}

public sealed class CitationDto
{
    public string SourceName { get; set; } =
        string.Empty;

    public string FileExtension { get; set; } =
        string.Empty;

    public string ViewerType { get; set; } =
        string.Empty;

    public int ChunkIndex { get; set; }

    public string CitationDisplay { get; set; } =
        string.Empty;

    public string ContentSnippet { get; set; } =
        string.Empty;

    public double Score { get; set; }

    public int? PageNumber { get; set; }

    public string? SheetName { get; set; }

    public string? CellAddress { get; set; }

    public string? HeadingPath { get; set; }

    public string? SearchText { get; set; }

    public Dictionary<string, string> CitationDetails { get; set; } =
        new(StringComparer.OrdinalIgnoreCase);
}
