#nullable enable

using Microsoft.AspNetCore.Http.Features;
using Microsoft.AspNetCore.StaticFiles;
using Npgsql;
using NpgsqlTypes;
using PgvectorVector = Pgvector.Vector;
using Syncfusion.DocumentChunking;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.EJ2.Spreadsheet;
using Syncfusion.Licensing;
using System.Linq;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;

using AgentRAGWithAzure.Models;

namespace AgentRAGWithAzure.Services;

public sealed class RagService
{
    private const int DefaultChunkCount = 10;

    private const int ChunkMultiplier = 4;

    private const int MinimumCandidateLimit = 50;

    private const double DefaultSimilarityThreshold = 0.0;

    private readonly HttpClient _geminiClient;

    private readonly string _geminiEmbeddingModel;

    private readonly string _geminiChatModel;

    private readonly int _embeddingDimensions;

    private readonly string _uploadFolder;

    private readonly NpgsqlDataSource _dataSource;

    private static readonly HashSet<string> SupportedExtensions =
        new(StringComparer.OrdinalIgnoreCase)
        {
            ".pdf",
            ".doc",
            ".docx",
            ".xls",
            ".xlsx"
        };

    public RagService(
        IConfiguration configuration,
        IWebHostEnvironment environment,
        NpgsqlDataSource dataSource)
    {
        _dataSource = dataSource;

        string geminiApiKey =
            configuration["Gemini:ApiKey"]
            ?? throw new InvalidOperationException(
                "Gemini:ApiKey is missing.");
        string geminiEndpoint =
            configuration["Gemini:Endpoint"]
            ?? "https://generativelanguage.googleapis.com/v1beta/";
        _geminiEmbeddingModel =
            configuration["Gemini:EmbeddingModel"]
            ?? "gemini-embedding-001";
        _geminiChatModel =
            configuration["Gemini:ChatModel"]
            ?? "gemini-3.8-flash";
        _embeddingDimensions =
            configuration.GetValue<int?>("Gemini:EmbeddingDimensions")
            ?? 1536;

        _geminiClient = new HttpClient
        {
            BaseAddress = new Uri(
                geminiEndpoint.TrimEnd('/') + "/"),
            Timeout = TimeSpan.FromMinutes(5)
        };
        _geminiClient.DefaultRequestHeaders.Add(
            "x-goog-api-key",
            geminiApiKey);

        // ------------------------------------------------------------
        // Use wwwroot/rag/uploads for uploaded source files
        // ------------------------------------------------------------

        string webRoot =
            environment.WebRootPath;

        if (string.IsNullOrWhiteSpace(webRoot))
        {
            webRoot =
                Path.Combine(
                    environment.ContentRootPath,
                    "wwwroot");
        }

        string ragRoot =
            Path.Combine(
                webRoot,
                "rag");

        _uploadFolder =
            Path.Combine(
                ragRoot,
                "uploads");

        Directory.CreateDirectory(_uploadFolder);
    }

    public bool IsSupportedFile(string fileName)
    {
        string extension =
            Path.GetExtension(fileName);

        return SupportedExtensions.Contains(extension);
    }

    public bool TryResolveUploadedFile(
        string fileName,
        out string fullPath,
        out string safeName,
        out string extension)
    {
        fullPath = string.Empty;
        safeName = string.Empty;
        extension = string.Empty;

        if (string.IsNullOrWhiteSpace(fileName))
        {
            return false;
        }

        safeName = Path.GetFileName(fileName);

        if (string.IsNullOrWhiteSpace(safeName) ||
            safeName.Contains("..", StringComparison.Ordinal))
        {
            return false;
        }

        extension = Path.GetExtension(safeName).ToLowerInvariant();

        if (!SupportedExtensions.Contains(extension))
        {
            return false;
        }

        fullPath = Path.GetFullPath(
            Path.Combine(_uploadFolder, safeName));

        string uploadRoot = Path.GetFullPath(_uploadFolder)
            .TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;

        if (!fullPath.StartsWith(uploadRoot, StringComparison.OrdinalIgnoreCase) ||
            !File.Exists(fullPath))
        {
            return false;
        }

        return true;
    }

    public async Task<bool> DeleteDocumentAsync(string fileName)
    {
        if (string.IsNullOrWhiteSpace(fileName))
        {
            return false;
        }

        string safeName = Path.GetFileName(fileName);
        if (string.IsNullOrWhiteSpace(safeName) ||
            safeName.Contains("..", StringComparison.Ordinal))
        {
            return false;
        }

        string extension = Path.GetExtension(safeName).ToLowerInvariant();
        if (!SupportedExtensions.Contains(extension))
        {
            return false;
        }

        string fullPath = Path.GetFullPath(Path.Combine(_uploadFolder, safeName));
        string uploadRoot = Path.GetFullPath(_uploadFolder)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(uploadRoot, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        await using NpgsqlConnection connection = await _dataSource.OpenConnectionAsync();
        await using NpgsqlCommand command = new(
            """
            DELETE FROM rag_chunks
            WHERE lower(source_name) = lower($1);
            """,
            connection);
        command.Parameters.AddWithValue(safeName);
        int deletedChunks = await command.ExecuteNonQueryAsync();

        bool fileExisted = File.Exists(fullPath);
        if (fileExisted)
        {
            File.Delete(fullPath);
        }

        return deletedChunks > 0 || fileExisted;
    }

    // ------------------------------------------------------------------------
    // Save, chunk, embed and store an uploaded document
    // ------------------------------------------------------------------------

    public async Task IndexUploadedFileAsync(
        IFormFile file,
        string safeFileName)
    {
        string filePath =
            Path.Combine(
                _uploadFolder,
                safeFileName);

        // Save the uploaded source file locally.
        await using (FileStream stream = new(
            filePath,
            FileMode.Create,
            FileAccess.Write,
            FileShare.None))
        {
            await file.CopyToAsync(stream);
        }

        // Chunk using Syncfusion Document Chunking.
        IReadOnlyList<IChunk> chunks =
            ChunkDocument(filePath);

        // Never report success when nothing was produced to store.
        chunks = chunks
            .Where(c => !string.IsNullOrWhiteSpace(c.Content))
            .ToList();

        if (chunks.Count == 0)
        {
            throw new InvalidOperationException(
                $"No text chunks could be extracted from '{safeFileName}', " +
                "so nothing was saved to the database.");
        }

        Console.WriteLine(
            $"Chunked '{safeFileName}' into {chunks.Count} chunks.");

        var records =
            new List<LocalChunkRecord>();

        foreach (IChunk chunk in chunks)
        {
            string searchableText =
                BuildSearchableText(
                    chunk,
                    safeFileName);

            float[] embedding =
                await GenerateEmbeddingAsync(
                    searchableText,
                    isQuery: false);

            string fileExtension =
                chunk.Metadata?.FileType
                ?? Path.GetExtension(safeFileName);

            records.Add(
                new LocalChunkRecord
                {
                    // The chunker's ChunkId may repeat across documents, which made
                    // ON CONFLICT (id) overwrite other rows. Make it unique per file.
                    Id = $"{Guid.NewGuid():N}",

                    // Use the filename consistently for delete and insert.
                    SourceName = safeFileName,

                    FileExtension = fileExtension,
                    ChunkIndex = chunk.ChunkIndex,
                    Content = chunk.Content,
                    Embedding = embedding,

                    MetadataAttributes =
                        chunk.Metadata?.Attributes
                        ?? new Dictionary<string, object>(),

                    CitationDisplay =
                        chunk.Citation?.DisplayText
                        ?? string.Empty,

                    CitationDetails =
                        chunk.Citation?.LocationDetails
                        ?? new Dictionary<string, object>()
                });
        }

        // Replace previous chunks for the same source file.
        await UpsertChunksAsync(
            records,
            safeFileName);

        Console.WriteLine(
            $"Stored {records.Count} chunks for '{safeFileName}'.");
    }

    // ------------------------------------------------------------------------
    // Build the text used to generate the embedding
    // ------------------------------------------------------------------------

    private static string BuildSearchableText(
        IChunk chunk,
        string sourceName)
    {
        var textBuilder =
            new StringBuilder();

        textBuilder.AppendLine(
            $"Source file: {sourceName}");

        if (chunk.Metadata?.Attributes != null)
        {
            foreach (
                KeyValuePair<string, object> attribute
                in chunk.Metadata.Attributes)
            {
                if (string.IsNullOrWhiteSpace(
                        attribute.Key) ||
                    attribute.Value == null)
                {
                    continue;
                }

                string? value =
                    ConvertMetadataValueToString(
                        attribute.Value);

                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                textBuilder.AppendLine(
                    $"{attribute.Key}: {value}");
            }
        }

        if (chunk.Citation != null)
        {
            textBuilder.AppendLine(
                $"Citation: {chunk.Citation.DisplayText}");

            if (chunk.Citation.LocationDetails != null)
            {
                foreach (
                    KeyValuePair<string, object> detail
                    in chunk.Citation.LocationDetails)
                {
                    if (string.IsNullOrWhiteSpace(
                            detail.Key) ||
                        detail.Value == null)
                    {
                        continue;
                    }

                    string? value =
                        ConvertMetadataValueToString(
                            detail.Value);

                    if (string.IsNullOrWhiteSpace(value))
                    {
                        continue;
                    }

                    textBuilder.AppendLine(
                        $"{detail.Key}: {value}");
                }
            }
        }

        textBuilder.AppendLine();
        textBuilder.Append(chunk.Content);

        return textBuilder.ToString();
    }

    // ------------------------------------------------------------------------
    // Chunk supported document formats only
    // ------------------------------------------------------------------------

    private static IReadOnlyList<IChunk> ChunkDocument(
        string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException(
                $"Input document not found: {filePath}");
        }

        var chunkingService =
            new ChunkingService();

        string extension =
            Path.GetExtension(filePath)
                .ToLowerInvariant();

        SourceChunkingOptions? sourceOptions =
            extension switch
            {
                ".doc" or ".docx" =>
                    new WordChunkingOptions
                    {
                        ChunkingMode =
                            WordChunkingMode.Auto
                    },

                ".pdf" =>
                    new PdfChunkingOptions
                    {
                        ChunkingMode =
                            PdfChunkingMode.Auto
                    },

                ".xlsx" or ".xls" =>
                    new ExcelChunkingOptions
                    {
                        ChunkingMode =
                            ExcelChunkingMode.Auto
                    },

                _ => null
            };

        if (sourceOptions == null)
        {
            throw new NotSupportedException(
                $"Unsupported document type for chunking: {extension}");
        }

        var options =
            new ChunkingOptions
            {
                MaxTokens = 450,
                OverlapTokens = 40,
                IncludeMetadata = true,
                IncludeCitation = true,
                SourceOptions = sourceOptions
            };

        IChunkingResult result =
            chunkingService.Chunk(
                filePath,
                options);

        return result.Chunks;
    }

    // ------------------------------------------------------------------------
    // Generate a Gemini embedding
    // ------------------------------------------------------------------------
    private async Task<float[]> GenerateEmbeddingAsync(
        string text,
        bool isQuery)
    {
        string model = Uri.EscapeDataString(_geminiEmbeddingModel);
        string requestUri = $"models/{model}:embedContent";
        var requestBody = new
        {
            model = $"models/{_geminiEmbeddingModel}",
            content = new
            {
                parts = new[]
                {
                    new { text }
                }
            },
            taskType = isQuery
                ? "RETRIEVAL_QUERY"
                : "RETRIEVAL_DOCUMENT",
            outputDimensionality = _embeddingDimensions
        };

        using HttpResponseMessage response =
            await _geminiClient.PostAsJsonAsync(
                requestUri,
                requestBody);
        string responseJson =
            await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Gemini embedding request failed ({(int)response.StatusCode}): " +
                responseJson);
        }

        using JsonDocument document =
            JsonDocument.Parse(responseJson);
        float[] values = document.RootElement
            .GetProperty("embedding")
            .GetProperty("values")
            .EnumerateArray()
            .Select(value => value.GetSingle())
            .ToArray();

        if (values.Length != _embeddingDimensions)
        {
            throw new InvalidOperationException(
                $"The Gemini embedding model '{_geminiEmbeddingModel}' " +
                $"returned {values.Length} dimensions, but {_embeddingDimensions} " +
                "dimensions were configured for PostgreSQL.");
        }
        return values;
    }

    // ------------------------------------------------------------------------
    // Replace all chunks belonging to the same uploaded document
    // ------------------------------------------------------------------------

    private async Task UpsertChunksAsync(
        IReadOnlyCollection<LocalChunkRecord> records,
        string sourceName)
    {
        await using NpgsqlConnection connection =
            await _dataSource.OpenConnectionAsync();

        await using NpgsqlTransaction transaction =
            await connection.BeginTransactionAsync();

        try
        {
            /*
             * This preserves the behavior of your previous JSON implementation.
             *
             * When the same filename is uploaded again:
             * 1. Delete its existing chunks.
             * 2. Insert the newly generated chunks.
             *
             * Both operations run in one transaction.
             */
            const string deleteSql =
                """
                DELETE FROM rag_chunks
                WHERE lower(source_name) = lower($1);
                """;

            await using (NpgsqlCommand deleteCommand = new(
                deleteSql,
                connection,
                transaction))
            {
                deleteCommand.Parameters.AddWithValue(
                    sourceName);

                await deleteCommand.ExecuteNonQueryAsync();
            }

            const string insertSql =
                """
                INSERT INTO rag_chunks
                (
                    id,
                    source_name,
                    file_extension,
                    chunk_index,
                    content,
                    embedding,
                    metadata_attributes,
                    citation_display,
                    citation_details
                )
                VALUES
                (
                    $1,
                    $2,
                    $3,
                    $4,
                    $5,
                    $6,
                    $7,
                    $8,
                    $9
                )
                ON CONFLICT (id)
                DO UPDATE SET
                    source_name =
                        EXCLUDED.source_name,

                    file_extension =
                        EXCLUDED.file_extension,

                    chunk_index =
                        EXCLUDED.chunk_index,

                    content =
                        EXCLUDED.content,

                    embedding =
                        EXCLUDED.embedding,

                    metadata_attributes =
                        EXCLUDED.metadata_attributes,

                    citation_display =
                        EXCLUDED.citation_display,

                    citation_details =
                        EXCLUDED.citation_details,

                    updated_at =
                        now();
                """;

            foreach (LocalChunkRecord record in records)
            {
                await using NpgsqlCommand insertCommand =
                    new(
                        insertSql,
                        connection,
                        transaction);

                insertCommand.Parameters.AddWithValue(
                    record.Id);

                insertCommand.Parameters.AddWithValue(
                    record.SourceName);

                insertCommand.Parameters.AddWithValue(
                    record.FileExtension);

                insertCommand.Parameters.AddWithValue(
                    record.ChunkIndex);

                insertCommand.Parameters.AddWithValue(
                    record.Content);

                insertCommand.Parameters.AddWithValue(
                    new PgvectorVector(record.Embedding));

                string metadataJson =
                    JsonSerializer.Serialize(
                        record.MetadataAttributes);

                insertCommand.Parameters.Add(
                    new NpgsqlParameter
                    {
                        NpgsqlDbType =
                            NpgsqlDbType.Jsonb,

                        Value =
                            metadataJson
                    });

                insertCommand.Parameters.AddWithValue(
                    record.CitationDisplay
                    ?? string.Empty);

                string citationDetailsJson =
                    JsonSerializer.Serialize(
                        record.CitationDetails);

                insertCommand.Parameters.Add(
                    new NpgsqlParameter
                    {
                        NpgsqlDbType =
                            NpgsqlDbType.Jsonb,

                        Value =
                            citationDetailsJson
                    });

                await insertCommand.ExecuteNonQueryAsync();
            }

            await using NpgsqlCommand countCommand = new(
                "SELECT count(*) FROM rag_chunks WHERE lower(source_name) = lower($1);",
                connection,
                transaction);
            countCommand.Parameters.AddWithValue(sourceName);
            long stored = Convert.ToInt64(await countCommand.ExecuteScalarAsync());
            Console.WriteLine($"rag_chunks now holds {stored} rows for '{sourceName}'.");

            if (stored != records.Count)
            {
                throw new InvalidOperationException(
                    $"Expected {records.Count} chunks in the database for '{sourceName}' but found {stored}.");
            }

            await transaction.CommitAsync();
        }
        catch
        {
            await transaction.RollbackAsync();

            throw;
        }
    }

    // ------------------------------------------------------------------------
    // Return distinct processed filenames from PostgreSQL
    // ------------------------------------------------------------------------

    public async Task<List<string>>
        GetProcessedFilesAsync()
    {
        const string sql =
            """
            SELECT DISTINCT source_name
            FROM rag_chunks
            WHERE source_name <> ''
            ORDER BY source_name;
            """;

        var files =
            new List<string>();

        await using NpgsqlConnection connection =
            await _dataSource.OpenConnectionAsync();

        await using NpgsqlCommand command =
            new(
                sql,
                connection);

        await using NpgsqlDataReader reader =
            await command.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            files.Add(
                reader.GetString(0));
        }

        return files;
    }

    // ------------------------------------------------------------------------
    // Generate the query embedding and search with pgvector cosine distance
    // ------------------------------------------------------------------------

    private async Task<List<LocalChunkRecord>>
        RetrieveRelevantChunksAsync(
            string question,
            int topK = DefaultChunkCount,
            double similarityThreshold =
                DefaultSimilarityThreshold)
    {
        if (topK <= 0)
        {
            return new List<LocalChunkRecord>();
        }

        float[] questionVector =
            await GenerateEmbeddingAsync(
                question,
                isQuery: true);

        Console.WriteLine(
            $"Question embedding dimensions: {questionVector.Length}");

        /*
         * Get more vector candidates than the final requested chunk count.
         * The similarity threshold is applied after candidate retrieval.
         */
        int candidateLimit =
            Math.Max(
                MinimumCandidateLimit,
                topK * ChunkMultiplier);

        /*
         * The pgvector <=> operator returns cosine distance.
         *
         * Similarity score:
         *
         *     1 - cosine distance
         *
         * The HNSW index uses vector_cosine_ops.
         */
        const string sql =
            """
            WITH nearest_results AS MATERIALIZED
            (
                SELECT
                    id,
                    source_name,
                    file_extension,
                    chunk_index,
                    content,
                    metadata_attributes,
                    citation_display,
                    citation_details,
                    embedding <=> $1 AS distance
                FROM rag_chunks
                ORDER BY embedding <=> $1
                LIMIT $2
            )
            SELECT
                id,
                source_name,
                file_extension,
                chunk_index,
                content,
                metadata_attributes,
                citation_display,
                citation_details,
                1 - distance AS score
            FROM nearest_results
            WHERE 1 - distance >= $3
            ORDER BY distance
            LIMIT $4;
            """;

        var records =
            new List<LocalChunkRecord>();

        await using NpgsqlConnection connection =
            await _dataSource.OpenConnectionAsync();

        await using NpgsqlTransaction transaction =
            await connection.BeginTransactionAsync();

        try
        {
            /*
             * The setting is local to this transaction.
             * It allows pgvector to continue scanning the HNSW index
             * when more candidates are needed.
             */
            const string settingsSql =
                """
                SET LOCAL hnsw.iterative_scan = relaxed_order;
                """;

            await using (NpgsqlCommand settingsCommand = new(
                settingsSql,
                connection,
                transaction))
            {
                await settingsCommand.ExecuteNonQueryAsync();
            }

            await using NpgsqlCommand command =
                new(
                    sql,
                    connection,
                    transaction);

            command.Parameters.AddWithValue(
                new PgvectorVector(questionVector));

            command.Parameters.AddWithValue(
                candidateLimit);

            command.Parameters.AddWithValue(
                similarityThreshold);

            command.Parameters.AddWithValue(
                topK);

            await using NpgsqlDataReader reader =
                await command.ExecuteReaderAsync();

            while (await reader.ReadAsync())
            {
                string metadataJson =
                    reader.GetString(5);

                string citationDetailsJson =
                    reader.GetString(7);

                Dictionary<string, object>? metadata =
                    JsonSerializer.Deserialize<
                        Dictionary<string, object>>(
                            metadataJson);

                Dictionary<string, object>? citationDetails =
                    JsonSerializer.Deserialize<
                        Dictionary<string, object>>(
                            citationDetailsJson);

                records.Add(
                    new LocalChunkRecord
                    {
                        Id =
                            reader.GetString(0),

                        SourceName =
                            reader.GetString(1),

                        FileExtension =
                            reader.GetString(2),

                        ChunkIndex =
                            reader.GetInt32(3),

                        Content =
                            reader.GetString(4),

                        MetadataAttributes =
                            metadata
                            ?? new Dictionary<string, object>(),

                        CitationDisplay =
                            reader.GetString(6),

                        CitationDetails =
                            citationDetails
                            ?? new Dictionary<string, object>(),

                        Score =
                            reader.GetDouble(8)
                    });
            }

            Console.WriteLine(
                $"Retrieved chunk count: {records.Count}");

            foreach (LocalChunkRecord record in records)
            {
                Console.WriteLine(
                    $"Source: {record.SourceName}, " +
                    $"Chunk: {record.ChunkIndex}, " +
                    $"Score: {record.Score:F4}");
            }

            await reader.CloseAsync();

            await transaction.CommitAsync();

            return records;
        }
        catch
        {
            await transaction.RollbackAsync();

            throw;
        }
    }

    // ------------------------------------------------------------------------
    // Retrieve relevant chunks and generate the final answer
    // ------------------------------------------------------------------------

    public async Task<ChatAnswerResult> AskQuestionAsync(
        string question)
    {
        List<LocalChunkRecord> chunks =
            await RetrieveRelevantChunksAsync(
                question,
                topK: 10);

        if (chunks.Count == 0)
        {
            return new ChatAnswerResult
            {
                Answer =
                    "I could not find relevant content " +
                    "in the uploaded knowledge files."
            };
        }

        string context =
            string.Join(
                Environment.NewLine +
                "--------------------" +
                Environment.NewLine,

                chunks.Select(
                    chunk =>
$"""
Source: {chunk.SourceName}
Chunk Index: {chunk.ChunkIndex}
Metadata:
{MetadataAttributesToString(chunk.MetadataAttributes)}
Citation: {chunk.CitationDisplay}
Citation Details:
{MetadataAttributesToString(chunk.CitationDetails)}
Score: {chunk.Score:F4}

{chunk.Content}
"""));

        // Prefer the single strongest match as the primary citation shown in UI.
        LocalChunkRecord primaryChunk = chunks[0];
        CitationDto primaryCitation = BuildCitation(primaryChunk);

        string prompt =
$"""
You are an enterprise RAG assistant.

Rules:
1. Use only the supplied context.
2. Do not invent information.
3. Keep the answer concise, scannable, and complete (Copilot-style).
4. Prefer short paragraphs and bullet points. Avoid long walls of text.
5. Mention the single best source once near the end (file name + page/sheet/heading if available).
6. Do NOT list multiple citation blocks, multiple "Click here" style links, or long source dumps.
7. Format lightly with:
   - Bullet points or numbered lists when helpful
   - Bold (**text**) for key terms
   - Clear line breaks between sections

Primary source to cite:
- File: {primaryCitation.SourceName}
- Location: {primaryCitation.CitationDisplay}
- Page: {(primaryCitation.PageNumber?.ToString() ?? "n/a")}
- Sheet/Cell: {(primaryCitation.SheetName ?? "n/a")} / {(primaryCitation.CellAddress ?? "n/a")}
- Heading: {(primaryCitation.HeadingPath ?? "n/a")}

Context:

{context}

Question:

{question}
""";

        string model = Uri.EscapeDataString(_geminiChatModel);
        string requestUri = $"models/{model}:generateContent";
        var requestBody = new
        {
            contents = new[]
            {
                new
                {
                    role = "user",
                    parts = new[]
                    {
                        new { text = prompt }
                    }
                }
            }
        };

        using HttpResponseMessage response =
            await _geminiClient.PostAsJsonAsync(
                requestUri,
                requestBody);
        string responseJson =
            await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode)
        {
            throw new InvalidOperationException(
                $"Gemini chat request failed ({(int)response.StatusCode}): " +
                responseJson);
        }

        using JsonDocument responseDocument =
            JsonDocument.Parse(responseJson);
        string answer = responseDocument.RootElement
            .GetProperty("candidates")[0]
            .GetProperty("content")
            .GetProperty("parts")[0]
            .GetProperty("text")
            .GetString()
            ?? string.Empty;

        return new ChatAnswerResult
        {
            Answer = answer,
            // Exactly one citation for precise jump + highlight in the viewer.
            Citations = new List<CitationDto> { primaryCitation }
        };
    }

    private static CitationDto BuildCitation(
        LocalChunkRecord chunk)
    {
        Dictionary<string, string> details =
            FlattenAttributes(chunk.CitationDetails);

        Dictionary<string, string> metadata =
            FlattenAttributes(chunk.MetadataAttributes);

        string extension =
            NormalizeExtension(
                chunk.FileExtension,
                chunk.SourceName);

        string snippet =
            BuildSnippet(chunk.Content);

        int? pageNumber =
            TryReadInt(details, metadata,
                "page", "pageNumber", "page_number", "pageno", "pageIndex", "page_index");

        // Some PDF chunkers store 0-based page indexes.
        if (pageNumber is null)
        {
            int? pageIndex =
                TryReadInt(details, metadata, "pageIndex", "page_index");

            if (pageIndex is >= 0)
            {
                pageNumber = pageIndex.Value + 1;
            }
        }

        string? sheetName =
            TryReadString(details, metadata,
                "sheet", "sheetName", "sheet_name", "worksheet", "worksheetName");

        string? cellAddress =
            TryReadString(details, metadata,
                "cell", "cellAddress", "cell_address", "address", "range", "cellRange");

        string? headingPath =
            TryReadString(details, metadata,
                "headingPath", "heading_path", "heading");

        // Prefer a stable phrase from the chunk BODY (not heading) so Word/PDF
        // search lands on the answering paragraph rather than only the heading.
        string? searchText = BuildSearchText(
            headingPath,
            snippet,
            chunk.Content);

        if (string.IsNullOrWhiteSpace(cellAddress) &&
            !string.IsNullOrWhiteSpace(chunk.CitationDisplay))
        {
            // Best-effort parse like "Sheet1!A1" or "A1".
            string display = chunk.CitationDisplay.Trim();
            int bang = display.LastIndexOf('!');

            if (bang > 0 && bang < display.Length - 1)
            {
                sheetName ??= display[..bang].Trim();
                cellAddress ??= display[(bang + 1)..].Trim();
            }
        }

        string viewerType = extension switch
        {
            ".pdf" => "pdf",
            ".doc" or ".docx" => "word",
            ".xls" or ".xlsx" => "excel",
            _ => "unknown"
        };

        string citationDisplay =
            string.IsNullOrWhiteSpace(chunk.CitationDisplay)
                ? BuildFallbackCitationDisplay(
                    chunk.SourceName,
                    pageNumber,
                    sheetName,
                    cellAddress,
                    headingPath)
                : chunk.CitationDisplay;

        return new CitationDto
        {
            SourceName = chunk.SourceName,
            FileExtension = extension,
            ViewerType = viewerType,
            ChunkIndex = chunk.ChunkIndex,
            CitationDisplay = citationDisplay,
            ContentSnippet = snippet,
            Score = chunk.Score,
            PageNumber = pageNumber,
            SheetName = sheetName,
            CellAddress = cellAddress,
            HeadingPath = headingPath,
            SearchText = searchText,
            CitationDetails = details
        };
    }

    private static string BuildFallbackCitationDisplay(
        string sourceName,
        int? pageNumber,
        string? sheetName,
        string? cellAddress,
        string? headingPath)
    {
        if (!string.IsNullOrWhiteSpace(headingPath))
        {
            return $"{sourceName} · {headingPath}";
        }

        if (!string.IsNullOrWhiteSpace(sheetName) ||
            !string.IsNullOrWhiteSpace(cellAddress))
        {
            string location =
                string.IsNullOrWhiteSpace(sheetName)
                    ? (cellAddress ?? string.Empty)
                    : string.IsNullOrWhiteSpace(cellAddress)
                        ? sheetName!
                        : $"{sheetName}!{cellAddress}";

            return $"{sourceName} · {location}";
        }

        if (pageNumber is > 0)
        {
            return $"{sourceName} · page {pageNumber}";
        }

        return sourceName;
    }

    private static string BuildSearchText(
        string? headingPath,
        string snippet,
        string content)
    {
        static string Clean(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            // Chunk extractors often join fields with | / — strip those so viewer search can match.
            string normalized = value
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace('|', ' ')
                .Replace('\\', ' ')
                .Replace('\u00a0', ' ');

            // Drop common list markers that rarely exist verbatim in DOCX text runs.
            normalized = normalized
                .Replace("•", " ")
                .Replace("·", " ");

            while (normalized.Contains("  ", StringComparison.Ordinal))
            {
                normalized = normalized.Replace("  ", " ", StringComparison.Ordinal);
            }

            return string.Join(
                " ",
                normalized.Split(
                    ' ',
                    StringSplitOptions.RemoveEmptyEntries |
                    StringSplitOptions.TrimEntries));
        }

        // Keep CONSECUTIVE body words. DocumentEditor/PDF find match literal spans;
        // stopword stripping produced non-contiguous phrases that never appear in DOCX text
        // (e.g. "Remote Hybrid Work Policy..." vs "Remote and Hybrid Work Policy...").
        static string? PhraseFrom(string text, int maxWords)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            string[] words = text.Split(
                ' ',
                StringSplitOptions.RemoveEmptyEntries |
                StringSplitOptions.TrimEntries);

            // Prefer the first run of real words (skip leading pure punctuation tokens).
            int start = 0;
            while (start < words.Length &&
                   words[start].All(ch => !char.IsLetterOrDigit(ch)))
            {
                start++;
            }

            if (start >= words.Length)
            {
                return null;
            }

            int take = Math.Min(Math.Max(2, maxWords), words.Length - start);
            return string.Join(" ", words.Skip(start).Take(take));
        }

        // Body/snippet first: Word navigation should land on the answering text,
        // not only the section heading (headings often appear earlier and mislead find).
        string body = Clean(
            !string.IsNullOrWhiteSpace(content)
                ? content
                : snippet);

        // Skip very short leading labels if a longer sentence follows later in the chunk.
        // Still keep contiguous words from wherever we start.
        string? bodyPhrase = PhraseFrom(body, 12);
        if (!string.IsNullOrWhiteSpace(bodyPhrase) && bodyPhrase.Length >= 12)
        {
            return bodyPhrase;
        }

        if (!string.IsNullOrWhiteSpace(bodyPhrase))
        {
            return bodyPhrase;
        }

        string? snippetPhrase = PhraseFrom(Clean(snippet), 10);
        if (!string.IsNullOrWhiteSpace(snippetPhrase))
        {
            return snippetPhrase;
        }

        // Heading is last-resort fallback only.
        string? headingPhrase = PhraseFrom(Clean(headingPath), 8);
        return headingPhrase ?? string.Empty;
    }

    private static Dictionary<string, string> FlattenAttributes(
        IReadOnlyDictionary<string, object>? attributes)
    {
        var map =
            new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);

        if (attributes == null)
        {
            return map;
        }

        foreach (KeyValuePair<string, object> pair in attributes)
        {
            if (string.IsNullOrWhiteSpace(pair.Key) ||
                pair.Value == null)
            {
                continue;
            }

            string? value =
                ConvertMetadataValueToString(pair.Value);

            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            map[pair.Key] = value;
        }

        return map;
    }

    private static string NormalizeExtension(
        string fileExtension,
        string sourceName)
    {
        string extension =
            string.IsNullOrWhiteSpace(fileExtension)
                ? Path.GetExtension(sourceName)
                : fileExtension;

        if (!extension.StartsWith(".", StringComparison.Ordinal))
        {
            extension = "." + extension;
        }

        return extension.ToLowerInvariant();
    }

    private static string BuildSnippet(string content)
    {
        if (string.IsNullOrWhiteSpace(content))
        {
            return string.Empty;
        }

        string normalized =
            string.Join(
                " ",
                content.Split(
                    (char[]?)null,
                    StringSplitOptions.RemoveEmptyEntries));

        return normalized.Length <= 180
            ? normalized
            : normalized[..180] + "…";
    }

    private static int? TryReadInt(
        Dictionary<string, string> details,
        Dictionary<string, string> metadata,
        params string[] keys)
    {
        string? raw =
            TryReadString(details, metadata, keys);

        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        if (int.TryParse(raw, out int value))
        {
            return value;
        }

        // Handle values like "Page 3".
        foreach (char separator in new[] { ' ', ':', '-' })
        {
            int idx = raw.LastIndexOf(separator);

            if (idx >= 0 &&
                int.TryParse(raw[(idx + 1)..].Trim(), out value))
            {
                return value;
            }
        }

        return null;
    }

    private static string? TryReadString(
        Dictionary<string, string> details,
        Dictionary<string, string> metadata,
        params string[] keys)
    {
        foreach (string key in keys)
        {
            if (details.TryGetValue(key, out string? fromDetails) &&
                !string.IsNullOrWhiteSpace(fromDetails))
            {
                return fromDetails.Trim();
            }

            if (metadata.TryGetValue(key, out string? fromMetadata) &&
                !string.IsNullOrWhiteSpace(fromMetadata))
            {
                return fromMetadata.Trim();
            }
        }

        return null;
    }

    private static string MetadataAttributesToString(
        IReadOnlyDictionary<string, object>? attributes)
    {
        if (attributes == null || attributes.Count == 0)
        {
            return "None";
        }

        var textBuilder =
            new StringBuilder();

        foreach (
            KeyValuePair<string, object> attribute
            in attributes)
        {
            if (string.IsNullOrWhiteSpace(
                    attribute.Key) ||
                attribute.Value == null)
            {
                continue;
            }

            string? value =
                ConvertMetadataValueToString(
                    attribute.Value);

            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            textBuilder.Append(attribute.Key);
            textBuilder.Append(": ");
            textBuilder.AppendLine(value);
        }

        return textBuilder.Length > 0
            ? textBuilder.ToString().TrimEnd()
            : "None";
    }

    // ------------------------------------------------------------------------
    // Convert normal objects and deserialized JSON values to strings
    // ------------------------------------------------------------------------

    private static string? ConvertMetadataValueToString(
        object value)
    {
        if (value is JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String =>
                    element.GetString(),

                JsonValueKind.Number =>
                    element.GetRawText(),

                JsonValueKind.True =>
                    bool.TrueString,

                JsonValueKind.False =>
                    bool.FalseString,

                JsonValueKind.Null =>
                    null,

                JsonValueKind.Undefined =>
                    null,

                _ =>
                    element.GetRawText()
            };
        }

        return value.ToString();
    }
}

// ============================================================================
// Models
// ============================================================================

