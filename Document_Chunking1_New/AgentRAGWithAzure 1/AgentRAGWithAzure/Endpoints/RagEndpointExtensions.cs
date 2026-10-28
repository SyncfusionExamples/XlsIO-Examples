using AgentRAGWithAzure.Models;
using AgentRAGWithAzure.Services;
using Microsoft.AspNetCore.StaticFiles;
using Syncfusion.EJ2.DocumentEditor;
using Syncfusion.EJ2.Spreadsheet;

namespace AgentRAGWithAzure.Endpoints;

public static class RagEndpointExtensions
{
    public static IEndpointRouteBuilder MapRagEndpoints(this IEndpointRouteBuilder endpoints)
    {// Return the list of processed files
// ------------------------------------------------------------

endpoints.MapGet(
    "/api/files",
    async (RagService ragService) =>
    {
        List<string> files =
            await ragService.GetProcessedFilesAsync();

        return Results.Ok(new
        {
            files
        });
    });

endpoints.MapDelete(
    "/api/documents/{fileName}",
    async (string fileName, RagService ragService) =>
    {
        bool removed = await ragService.DeleteDocumentAsync(fileName);
        if (!removed)
        {
            return Results.NotFound(new { error = "Document not found." });
        }

        return Results.Ok(new { deleted = fileName });
    });

// ------------------------------------------------------------
// Upload and index documents
// ------------------------------------------------------------

endpoints.MapPost(
    "/api/upload",
    async (
        HttpRequest request,
        RagService ragService) =>
    {
        if (!request.HasFormContentType)
        {
            return Results.BadRequest(
                new UploadErrorResponse
                {
                    Error = "Invalid upload request."
                });
        }

        IFormCollection form =
            await request.ReadFormAsync();

        IFormFileCollection files =
            form.Files;

        if (files.Count == 0)
        {
            return Results.BadRequest(
                new UploadErrorResponse
                {
                    Error =
                        "Please upload at least one supported file."
                });
        }

        var uploadedFiles =
            new List<string>();

        foreach (IFormFile file in files)
        {
            if (!ragService.IsSupportedFile(file.FileName))
            {
                continue;
            }

            string safeFileName =
                Path.GetFileName(file.FileName);

            await ragService.IndexUploadedFileAsync(
                file,
                safeFileName);

            uploadedFiles.Add(safeFileName);
        }

        if (uploadedFiles.Count == 0)
        {
            return Results.BadRequest(
                new UploadErrorResponse
                {
                    Error =
                        "No supported files were uploaded. " +
                        "Supported formats: .pdf, .doc, .docx, " +
                        ".xls, .xlsx"
                });
        }

        return Results.Ok(
            new UploadResponse
            {
                Files = uploadedFiles
            });
    });

// ------------------------------------------------------------
// Ask questions using retrieved PostgreSQL chunks
// ------------------------------------------------------------

endpoints.MapPost(
    "/api/chat",
    async (
        ChatRequest request,
        RagService ragService) =>
    {
        if (string.IsNullOrWhiteSpace(request.Question))
        {
            return Results.BadRequest(
                new ChatErrorResponse
                {
                    Error = "Question is required."
                });
        }

        try
        {
            ChatAnswerResult result =
                await ragService.AskQuestionAsync(
                    request.Question);

            return Results.Ok(
                new ChatResponse
                {
                    Answer = result.Answer,
                    Citations = result.Citations
                });
        }
        catch (Exception ex)
        {
            return Results.Problem(
                detail: ex.Message,
                title: "Could not answer that question.");
        }
    });

// ------------------------------------------------------------
// Document open APIs for PDF / Word / Excel viewers
// ------------------------------------------------------------

endpoints.MapGet(
    "/api/documents/{fileName}/meta",
    (string fileName, RagService ragService) =>
    {
        if (!ragService.TryResolveUploadedFile(
                fileName,
                out string fullPath,
                out string safeName,
                out string extension))
        {
            return Results.NotFound(new
            {
                error = "Document not found."
            });
        }

        return Results.Ok(new
        {
            fileName = safeName,
            extension,
            exists = File.Exists(fullPath),
            size = new FileInfo(fullPath).Length
        });
    });

endpoints.MapGet(
    "/api/documents/{fileName}/stream",
    (string fileName, RagService ragService) =>
    {
        if (!ragService.TryResolveUploadedFile(
                fileName,
                out string fullPath,
                out string safeName,
                out string extension))
        {
            return Results.NotFound("Document not found.");
        }

        string contentType =
            GetContentType(extension);

        FileStream stream = new(
            fullPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read);

        return Results.File(
            stream,
            contentType,
            fileDownloadName: safeName,
            enableRangeProcessing: true);
    });

endpoints.MapGet(
    "/api/documents/{fileName}/base64",
    async (string fileName, RagService ragService) =>
    {
        if (!ragService.TryResolveUploadedFile(
                fileName,
                out string fullPath,
                out string safeName,
                out string extension))
        {
            return Results.NotFound(new
            {
                error = "Document not found."
            });
        }

        byte[] bytes =
            await File.ReadAllBytesAsync(fullPath);

        return Results.Ok(new
        {
            fileName = safeName,
            extension,
            contentType = GetContentType(extension),
            base64 = Convert.ToBase64String(bytes)
        });
    });

endpoints.MapGet(
    "/api/documents/{fileName}/sfdt",
    (string fileName, RagService ragService) =>
    {
        if (!ragService.TryResolveUploadedFile(
                fileName,
                out string fullPath,
                out string safeName,
                out string extension))
        {
            return Results.NotFound(new
            {
                error = "Document not found."
            });
        }

        if (extension is not (".doc" or ".docx" or ".dot" or ".dotx" or ".dotm" or ".rtf"))
        {
            return Results.BadRequest(new
            {
                error = "SFDT conversion is only supported for Word documents."
            });
        }

        try
        {
            string sfdt =
                ConvertWordFileToSfdt(fullPath, extension);

            return Results.Content(sfdt, "application/json");
        }
        catch (Exception ex)
        {
            return Results.Problem(
                detail: ex.Message,
                title: "Failed to convert Word document to SFDT.");
        }
    });

// ------------------------------------------------------------
// Spreadsheet open (EJ2 openUrl expects multipart file → workbook JSON)
// ------------------------------------------------------------

endpoints.MapPost(
    "/api/spreadsheet/open",
    async (HttpRequest request) =>
    {
        if (!request.HasFormContentType)
        {
            return Results.BadRequest(new
            {
                error = "Expected multipart form data with an Excel file."
            });
        }

        IFormCollection form =
            await request.ReadFormAsync();

        IFormFile? file =
            form.Files.Count > 0
                ? form.Files[0]
                : null;

        if (file is null || file.Length <= 0)
        {
            return Results.BadRequest(new
            {
                error = "Excel file is required."
            });
        }

        try
        {
            OpenRequest open = new()
            {
                File = file
            };

            if (form.TryGetValue(
                    "chunkPayload",
                    out Microsoft.Extensions.Primitives.StringValues chunkPayload) &&
                !string.IsNullOrWhiteSpace(chunkPayload))
            {
                open.ChunkPayload = chunkPayload.ToString();
            }

            string workbookJson =
                Workbook.Open(open);

            return Results.Content(
                workbookJson,
                "application/json");
        }
        catch (Exception ex)
        {
            return Results.Problem(
                detail: ex.Message,
                title: "Failed to open spreadsheet.");
        }
    });

endpoints.MapGet(
    "/api/documents/{fileName}/workbook",
    (string fileName, RagService ragService) =>
    {
        if (!ragService.TryResolveUploadedFile(
                fileName,
                out string fullPath,
                out string safeName,
                out string extension))
        {
            return Results.NotFound(new
            {
                error = "Document not found."
            });
        }

        if (extension is not (".xls" or ".xlsx" or ".xlsm" or ".xlsb" or ".csv"))
        {
            return Results.BadRequest(new
            {
                error = "Workbook conversion is only supported for Excel files."
            });
        }

        try
        {
            string workbookJson =
                ConvertExcelFileToWorkbookJson(fullPath, safeName);

            return Results.Content(
                workbookJson,
                "application/json");
        }
        catch (Exception ex)
        {
            return Results.Problem(
                detail: ex.Message,
                title: "Failed to convert Excel document for Spreadsheet.");
        }
    });


        return endpoints;
    }

static string GetContentType(string extension)
{
    var provider = new FileExtensionContentTypeProvider();

    if (provider.TryGetContentType("file" + extension, out string? contentType) &&
        !string.IsNullOrWhiteSpace(contentType))
    {
        return contentType;
    }

    return extension.ToLowerInvariant() switch
    {
        ".pdf" => "application/pdf",
        ".doc" => "application/msword",
        ".docx" =>
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        ".xls" => "application/vnd.ms-excel",
        ".xlsx" =>
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        _ => "application/octet-stream"
    };
}

static string ConvertWordFileToSfdt(
    string filePath,
    string extension)
{
    using FileStream stream = new(
        filePath,
        FileMode.Open,
        FileAccess.Read,
        FileShare.Read);

    // EJ2 DocumentEditor SFDT import path (same approach as DocumentEditorController.Import).
    // WordDocument does not implement IDisposable; dispose explicitly when available.
    WordDocument ej2Document =
        WordDocument.Load(stream, GetWordFormatType(extension));

    try
    {
        string json =
            Newtonsoft.Json.JsonConvert.SerializeObject(ej2Document);

        return json;
    }
    finally
    {
        ej2Document.Dispose();
    }
}

static FormatType GetWordFormatType(string extension)
{
    return extension.ToLowerInvariant() switch
    {
        ".doc" or ".dot" or ".rtf" => FormatType.Doc,
        ".docx" or ".dotx" or ".dotm" => FormatType.Docx,
        _ => FormatType.Docx
    };
}

static string ConvertExcelFileToWorkbookJson(
    string filePath,
    string fileName)
{
    using FileStream stream = new(
        filePath,
        FileMode.Open,
        FileAccess.Read,
        FileShare.Read);

    // Workbook.Open expects an IFormFile; wrap the local upload stream.
    IFormFile formFile = new FormFile(
        stream,
        baseStreamOffset: 0,
        length: stream.Length,
        name: "file",
        fileName: fileName)
    {
        Headers = new HeaderDictionary(),
        ContentType = GetContentType(
            Path.GetExtension(fileName))
    };

    OpenRequest open = new()
    {
        File = formFile
    };

    return Workbook.Open(open);
}

// ============================================================================

}
