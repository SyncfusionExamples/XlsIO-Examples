#nullable enable

using AgentRAGWithAzure.Services;

using System.Net;
using System.Text;
using System.Text.Json;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using Newtonsoft.Json;
using Syncfusion.EJ2.PdfViewer;

namespace AgentRAGWithAzure.Controllers;

/// <summary>
/// Server-backed Syncfusion PDF Viewer endpoints (Syncfusion.EJ2.PdfViewer.AspNet.Core).
/// Loads RAG upload PDFs via <see cref="RagService"/> and processes pages on the server.
/// </summary>
[ApiController]
[Route("api/[controller]")]
[IgnoreAntiforgeryToken]
public sealed class PdfViewerController : ControllerBase
{
    private readonly IWebHostEnvironment _environment;
    private readonly IMemoryCache _cache;
    private readonly RagService _ragService;
    private readonly ILogger<PdfViewerController> _logger;

    public PdfViewerController(
        IWebHostEnvironment environment,
        IMemoryCache cache,
        RagService ragService,
        ILogger<PdfViewerController> logger)
    {
        _environment = environment;
        _cache = cache;
        _ragService = ragService;
        _logger = logger;
    }

    [HttpPost("Load")]
    [RequestSizeLimit(200_000_000)]
    public async Task<IActionResult> Load()
    {
        try
        {
            Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
            EnsureAction(jsonObject, "Load");
            EnsureViewerSessionKeys(jsonObject);

            // Treat short non-base64 tokens as file names even if isFileName is missing/false.
            // EJ2 "Invalid PDF file type..." often happens when a .pdf name is parsed as base64.
            if (jsonObject.TryGetValue("document", out string? docToken) &&
                LooksLikeFileName(docToken))
            {
                jsonObject["isFileName"] = "true";
            }

            PdfRenderer pdfviewer = new(_cache);
            MemoryStream? loadStream = null;

            _logger.LogInformation(
                "PdfViewer Load keys=[{Keys}] document={Document} isFileName={IsFileName} action={Action} documentId={DocumentId}",
                string.Join(",", jsonObject.Keys),
                Truncate(GetValue(jsonObject, "document"), 80),
                GetValue(jsonObject, "isFileName"),
                GetValue(jsonObject, "action"),
                GetValue(jsonObject, "documentId"));

            if (jsonObject.TryGetValue("document", out string? document) &&
                !string.IsNullOrWhiteSpace(document))
            {
                bool isFileName = IsTrue(GetValue(jsonObject, "isFileName")) || LooksLikeFileName(document);

                if (isFileName)
                {
                    if (!TryOpenDocumentStream(document, out MemoryStream opened, out string? error))
                    {
                        _logger.LogWarning("PdfViewer Load file not found: {Document} ({Error})", document, error);
                        // Return structured JSON so the client does not surface a generic invalid-PDF dialog.
                        return StatusCode(
                            StatusCodes.Status404NotFound,
                            new { error = error ?? (document + " is not found"), document });
                    }

                    // Keep an independent, non-disposed stream in memory cache for later RenderPdfPages.
                    loadStream = new MemoryStream(opened.ToArray(), writable: false);
                    opened.Dispose();
                }
                else
                {
                    try
                    {
                        string base64 = document;
                        int comma = base64.IndexOf(',');
                        if (comma >= 0 && base64.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                        {
                            base64 = base64[(comma + 1)..];
                        }

                        byte[] bytes = Convert.FromBase64String(base64);
                        loadStream = new MemoryStream(bytes, writable: false);
                    }
                    catch (FormatException ex)
                    {
                        _logger.LogWarning(ex, "PdfViewer Load base64 decode failed; document looked like payload but was invalid");
                        return StatusCode(
                            StatusCodes.Status400BadRequest,
                            new { error = "Invalid PDF payload (base64 decode failed)." });
                    }
                }
            }

            if (loadStream is null || loadStream.Length == 0)
            {
                _logger.LogWarning("PdfViewer Load empty stream / missing document.");
                return StatusCode(
                    StatusCodes.Status400BadRequest,
                    new { error = "PDF document stream is empty or document was not provided." });
            }

            if (loadStream.CanSeek)
            {
                loadStream.Position = 0;
            }

            // Do NOT dispose loadStream — PdfRenderer keeps it in IMemoryCache for page renders.
            object jsonResult = pdfviewer.Load(loadStream, jsonObject);
            string payload = jsonResult is string s
                ? s
                : JsonConvert.SerializeObject(jsonResult);

            // PdfRenderer sometimes returns a plain error string with HTTP 200.
            // That triggers the EJ2 "Invalid PDF file type..." dialog — fail hard instead.
            if (IsPdfRendererErrorPayload(payload))
            {
                _logger.LogWarning("PdfViewer Load returned error payload: {Error}", Truncate(payload, 200));
                return StatusCode(
                    StatusCodes.Status400BadRequest,
                    new { error = payload.Trim('"') });
            }

            _logger.LogInformation("PdfViewer Load ok length={Length}", payload.Length);
            return JsonContent(payload);
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "PdfViewer Load failed");
            return StatusCode(StatusCodes.Status500InternalServerError, new { error = ex.Message });
        }
    }

    [HttpPost("ValidatePassword")]
    [RequestSizeLimit(200_000_000)]
    public async Task<IActionResult> ValidatePassword()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "ValidatePassword");

        PdfRenderer pdfviewer = new(_cache);
        MemoryStream stream = new();

        if (jsonObject.TryGetValue("document", out string? document) &&
            !string.IsNullOrWhiteSpace(document))
        {
            if (IsTrue(GetValue(jsonObject, "isFileName")))
            {
                if (!TryOpenDocumentStream(document, out stream, out string? error))
                {
                    return JsonContent(error ?? (document + " is not found"));
                }
            }
            else
            {
                byte[] bytes = Convert.FromBase64String(document);
                stream = new MemoryStream(bytes);
            }
        }

        string? password = GetValue(jsonObject, "password");
        if (stream.CanSeek)
        {
            stream.Position = 0;
        }

        object result = pdfviewer.Load(stream, password);
        return JsonContent(JsonConvert.SerializeObject(result));
    }

    [HttpPost("RenderPdfPages")]
    public async Task<IActionResult> RenderPdfPages()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "RenderPdfPages");
        PdfRenderer pdfviewer = new(_cache);
        object jsonResult = pdfviewer.GetPage(jsonObject);
        return JsonContent(JsonConvert.SerializeObject(jsonResult));
    }

    [HttpPost("RenderPdfTexts")]
    public async Task<IActionResult> RenderPdfTexts()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "RenderPdfTexts");
        PdfRenderer pdfviewer = new(_cache);
        object jsonResult = pdfviewer.GetDocumentText(jsonObject);
        return JsonContent(JsonConvert.SerializeObject(jsonResult));
    }

    [HttpPost("Unload")]
    public async Task<IActionResult> Unload()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "Unload");
        PdfRenderer pdfviewer = new(_cache);
        pdfviewer.ClearCache(jsonObject);
        return Content("Document cache is cleared");
    }

    [HttpPost("RenderThumbnailImages")]
    public async Task<IActionResult> RenderThumbnailImages()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "RenderThumbnailImages");
        PdfRenderer pdfviewer = new(_cache);
        object result = pdfviewer.GetThumbnailImages(jsonObject);
        return JsonContent(JsonConvert.SerializeObject(result));
    }

    [HttpPost("Bookmarks")]
    public async Task<IActionResult> Bookmarks()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "Bookmarks");
        PdfRenderer pdfviewer = new(_cache);
        object jsonResult = pdfviewer.GetBookmarks(jsonObject);
        return JsonContent(JsonConvert.SerializeObject(jsonResult));
    }

    [HttpPost("RenderAnnotationComments")]
    public async Task<IActionResult> RenderAnnotationComments()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "RenderAnnotationComments");
        PdfRenderer pdfviewer = new(_cache);
        object jsonResult = pdfviewer.GetAnnotationComments(jsonObject);
        return JsonContent(JsonConvert.SerializeObject(jsonResult));
    }

    [HttpPost("ExportAnnotations")]
    public async Task<IActionResult> ExportAnnotations()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "ExportAnnotations");
        PdfRenderer pdfviewer = new(_cache);
        string jsonResult = pdfviewer.ExportAnnotation(jsonObject);
        return Content(jsonResult);
    }

    [HttpPost("ImportAnnotations")]
    public async Task<IActionResult> ImportAnnotations()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "ImportAnnotations");
        PdfRenderer pdfviewer = new(_cache);

        if (jsonObject.TryGetValue("fileName", out string? fileName) &&
            !string.IsNullOrWhiteSpace(fileName))
        {
            string? documentPath = GetDocumentPath(fileName);
            if (string.IsNullOrEmpty(documentPath))
            {
                return Content(fileName + " is not found");
            }

            string text = System.IO.File.ReadAllText(documentPath);
            return Content(text);
        }

        if (jsonObject.TryGetValue("importedData", out string? importedData) &&
            !string.IsNullOrWhiteSpace(importedData))
        {
            string extension = Path.GetExtension(importedData);
            if (string.Equals(extension, ".xfdf", StringComparison.OrdinalIgnoreCase))
            {
                string? documentPath = GetDocumentPath(importedData);
                if (string.IsNullOrEmpty(documentPath))
                {
                    return Content(importedData + " is not found");
                }

                byte[] bytes = System.IO.File.ReadAllBytes(documentPath);
                jsonObject["importedData"] = Convert.ToBase64String(bytes);
            }
        }

        object jsonResult = pdfviewer.ImportAnnotation(jsonObject);
        return JsonContent(JsonConvert.SerializeObject(jsonResult));
    }

    [HttpPost("Download")]
    public async Task<IActionResult> Download()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "Download");
        PdfRenderer pdfviewer = new(_cache);
        string documentBase = pdfviewer.GetDocumentAsBase64(jsonObject);
        return Content(documentBase);
    }

    [HttpPost("PrintImages")]
    public async Task<IActionResult> PrintImages()
    {
        Dictionary<string, string> jsonObject = await ReadRequestDictionaryAsync();
        EnsureAction(jsonObject, "PrintImages");
        PdfRenderer pdfviewer = new(_cache);
        object pageImage = pdfviewer.GetPrintImage(jsonObject);
        return JsonContent(JsonConvert.SerializeObject(pageImage));
    }

    /// <summary>
    /// Read body as Newtonsoft JObject first (EJ2 client payload), fall back to form fields / System.Text.Json.
    /// PdfRenderer requires Dictionary<string,string> with an "action" key.
    /// </summary>
    private async Task<Dictionary<string, string>> ReadRequestDictionaryAsync()
    {
        Dictionary<string, string> result = new(StringComparer.OrdinalIgnoreCase);

        // EJ2 often posts application/x-www-form-urlencoded with jsonObject=...
        if (Request.HasFormContentType)
        {
            IFormCollection form = await Request.ReadFormAsync();
            string? nested =
                form["jsonObject"].FirstOrDefault()
                ?? form["jsonData"].FirstOrDefault()
                ?? form["data"].FirstOrDefault();

            if (!string.IsNullOrWhiteSpace(nested) &&
                TryParseJsonDictionary(nested, result))
            {
                return result;
            }

            foreach (string key in form.Keys)
            {
                string? value = form[key].FirstOrDefault();
                if (!string.IsNullOrEmpty(value))
                {
                    result[key] = value;
                }
            }

            if (result.Count > 0)
            {
                return result;
            }
        }

        Request.EnableBuffering();
        if (Request.Body.CanSeek)
        {
            Request.Body.Position = 0;
        }

        using StreamReader reader = new(Request.Body, Encoding.UTF8, detectEncodingFromByteOrderMarks: false, leaveOpen: true);
        string body = await reader.ReadToEndAsync();
        if (Request.Body.CanSeek)
        {
            Request.Body.Position = 0;
        }

        if (string.IsNullOrWhiteSpace(body))
        {
            return result;
        }

        // form-urlencoded body may have been buffered as string (no HasFormContentType in some hosts)
        if (body.Contains('=') && !body.TrimStart().StartsWith('{'))
        {
            string? nested = ExtractFormField(body, "jsonObject")
                ?? ExtractFormField(body, "jsonData")
                ?? ExtractFormField(body, "data");
            if (!string.IsNullOrWhiteSpace(nested) &&
                TryParseJsonDictionary(Uri.UnescapeDataString(nested.Replace('+', ' ')), result))
            {
                return result;
            }
        }

        if (TryParseJsonDictionary(body, result))
        {
            return result;
        }

        _logger.LogWarning(
            "Failed to parse PdfViewer request body (len={Length}, contentType={ContentType})",
            body.Length,
            Request.ContentType);
        return result;
    }

    private static string? ExtractFormField(string body, string fieldName)
    {
        string prefix = fieldName + "=";
        int start = body.IndexOf(prefix, StringComparison.OrdinalIgnoreCase);
        if (start < 0)
        {
            return null;
        }

        start += prefix.Length;
        int end = body.IndexOf('&', start);
        return end < 0 ? body[start..] : body[start..end];
    }

    private static bool TryParseJsonDictionary(string body, Dictionary<string, string> result)
    {
        if (string.IsNullOrWhiteSpace(body))
        {
            return false;
        }

        try
        {
            Newtonsoft.Json.Linq.JObject? jObject =
                JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(body);
            if (jObject is null)
            {
                return false;
            }

            foreach (Newtonsoft.Json.Linq.JProperty property in jObject.Properties())
            {
                if (property.Value is null ||
                    property.Value.Type is Newtonsoft.Json.Linq.JTokenType.Null or Newtonsoft.Json.Linq.JTokenType.Undefined)
                {
                    continue;
                }

                if (property.Value.Type == Newtonsoft.Json.Linq.JTokenType.Boolean)
                {
                    bool flag = property.Value.ToObject<bool>();
                    result[property.Name] = flag ? "true" : "false";
                }
                else if (property.Value.Type is Newtonsoft.Json.Linq.JTokenType.Object
                         or Newtonsoft.Json.Linq.JTokenType.Array)
                {
                    result[property.Name] = property.Value.ToString(Formatting.None);
                }
                else
                {
                    string text = property.Value.ToString();
                    if (!string.IsNullOrEmpty(text))
                    {
                        result[property.Name] = text;
                    }
                }
            }

            return result.Count > 0;
        }
        catch
        {
            // fall through to System.Text.Json
        }

        try
        {
            using JsonDocument doc = JsonDocument.Parse(body);
            if (doc.RootElement.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            foreach (JsonProperty property in doc.RootElement.EnumerateObject())
            {
                if (property.Value.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null)
                {
                    continue;
                }

                result[property.Name] = property.Value.ValueKind switch
                {
                    JsonValueKind.String => property.Value.GetString() ?? string.Empty,
                    JsonValueKind.True => "true",
                    JsonValueKind.False => "false",
                    JsonValueKind.Number => property.Value.GetRawText(),
                    _ => property.Value.GetRawText()
                };
            }

            return result.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    private static void EnsureAction(Dictionary<string, string> jsonObject, string defaultAction)
    {
        if (!jsonObject.TryGetValue("action", out string? action) ||
            string.IsNullOrWhiteSpace(action))
        {
            jsonObject["action"] = defaultAction;
        }
    }

    /// <summary>
    /// PdfRenderer caches the document under documentId/uniqueId/hashId.
    /// Without these keys Load can return "Cannot access a closed Stream" / cache errors.
    /// Prefer unique ids — reusing a stable filename-based id after Unload causes
    /// "Document stream does not exist in the cache" and the EJ2 invalid-PDF dialog.
    /// </summary>
    private static void EnsureViewerSessionKeys(Dictionary<string, string> jsonObject)
    {
        string fresh = "pdf_" + Guid.NewGuid().ToString("N");

        string? documentId = GetValue(jsonObject, "documentId");
        string? uniqueId = GetValue(jsonObject, "uniqueId");
        string? hashId = GetValue(jsonObject, "hashId");

        // Replace empty OR our old stable filename seed (doc_*.pdf) which collides after unload.
        bool needsFreshDocumentId =
            string.IsNullOrWhiteSpace(documentId)
            || documentId.StartsWith("doc_", StringComparison.OrdinalIgnoreCase);

        if (needsFreshDocumentId)
        {
            jsonObject["documentId"] = string.IsNullOrWhiteSpace(uniqueId) ? fresh : uniqueId;
        }

        if (string.IsNullOrWhiteSpace(uniqueId))
        {
            jsonObject["uniqueId"] = jsonObject["documentId"];
        }

        if (string.IsNullOrWhiteSpace(hashId)
            || hashId.StartsWith("doc_", StringComparison.OrdinalIgnoreCase))
        {
            jsonObject["hashId"] = jsonObject["documentId"];
        }

        if (!jsonObject.TryGetValue("elementId", out string? elementId) ||
            string.IsNullOrWhiteSpace(elementId))
        {
            jsonObject["elementId"] = "pdfViewer";
        }
    }

    private ContentResult JsonContent(string payload) =>
        Content(payload ?? string.Empty, "application/json", Encoding.UTF8);

    private static string? GetValue(Dictionary<string, string> jsonObject, string key) =>
        jsonObject.TryGetValue(key, out string? value) ? value : null;

    private static bool IsTrue(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        return bool.TryParse(value, out bool parsed) && parsed
               || string.Equals(value, "1", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsPdfRendererErrorPayload(string? payload)
    {
        if (string.IsNullOrWhiteSpace(payload))
        {
            return true;
        }

        string text = payload.Trim().Trim('"');
        return text.Contains("was not present in the dictionary", StringComparison.OrdinalIgnoreCase)
               || text.Contains("does not exist in the cache", StringComparison.OrdinalIgnoreCase)
               || text.Contains("closed Stream", StringComparison.OrdinalIgnoreCase)
               || text.Contains("Object reference not set", StringComparison.OrdinalIgnoreCase)
               || string.Equals(text, "Document stream does not exist in the cache", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// True for short path-like tokens (e.g. "report.pdf") that must never be treated as base64.
    /// </summary>
    private static bool LooksLikeFileName(string? document)
    {
        if (string.IsNullOrWhiteSpace(document))
        {
            return false;
        }

        string token = document.Trim();
        if (token.Length is 0 or > 260)
        {
            return false;
        }

        if (token.StartsWith("data:", StringComparison.OrdinalIgnoreCase) ||
            token.Contains(','))
        {
            return false;
        }

        // Real base64 payloads are long and do not look like file paths.
        if (token.Contains('/') || token.Contains('\\') || token.Contains('.'))
        {
            string name = Path.GetFileName(token.Split('?', 2)[0].Split('#', 2)[0]);
            return !string.IsNullOrWhiteSpace(name) &&
                   name.IndexOf('.') >= 0 &&
                   name.Length <= 200;
        }

        return false;
    }

    private static string Truncate(string? value, int max)
    {
        if (string.IsNullOrEmpty(value))
        {
            return string.Empty;
        }

        return value.Length <= max ? value : value[..max] + "…";
    }

    private bool TryOpenDocumentStream(
        string document,
        out MemoryStream stream,
        out string? error)
    {
        stream = new MemoryStream();
        error = null;

        if (string.IsNullOrWhiteSpace(document))
        {
            error = "Document name is empty.";
            return false;
        }

        string cleaned = document.Split('?', 2)[0].Split('#', 2)[0].Trim();
        string fileNameOnly = Path.GetFileName(cleaned);

        // Prefer RAG uploads (filename only).
        if (!string.IsNullOrWhiteSpace(fileNameOnly) &&
            _ragService.TryResolveUploadedFile(
                fileNameOnly,
                out string fullPath,
                out _,
                out string extension) &&
            string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase))
        {
            byte[] bytes = System.IO.File.ReadAllBytes(fullPath);
            stream = new MemoryStream(bytes);
            _logger.LogInformation(
                "PdfViewer resolved RAG file {File} -> {Path} ({Bytes} bytes)",
                fileNameOnly,
                fullPath,
                bytes.Length);
            return true;
        }

        // Also try the raw document token (may already be a bare name).
        if (!string.Equals(cleaned, fileNameOnly, StringComparison.OrdinalIgnoreCase) &&
            _ragService.TryResolveUploadedFile(
                cleaned,
                out string fullPath2,
                out _,
                out string extension2) &&
            string.Equals(extension2, ".pdf", StringComparison.OrdinalIgnoreCase))
        {
            byte[] bytes = System.IO.File.ReadAllBytes(fullPath2);
            stream = new MemoryStream(bytes);
            return true;
        }

        string? documentPath = GetDocumentPath(cleaned);
        if (string.IsNullOrEmpty(documentPath) &&
            !string.Equals(cleaned, fileNameOnly, StringComparison.OrdinalIgnoreCase))
        {
            documentPath = GetDocumentPath(fileNameOnly);
        }

        if (!string.IsNullOrEmpty(documentPath))
        {
            byte[] bytes = System.IO.File.ReadAllBytes(documentPath);
            stream = new MemoryStream(bytes);
            _logger.LogInformation(
                "PdfViewer resolved path {Path} ({Bytes} bytes)",
                documentPath,
                bytes.Length);
            return true;
        }

        string scheme = cleaned.Split(new[] { "://" }, StringSplitOptions.None)[0];
        if (string.Equals(scheme, "http", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scheme, "https", StringComparison.OrdinalIgnoreCase))
        {
#pragma warning disable SYSLIB0014
            using WebClient webClient = new();
            byte[] pdfDoc = webClient.DownloadData(cleaned);
#pragma warning restore SYSLIB0014
            stream = new MemoryStream(pdfDoc);
            return true;
        }

        error = cleaned + " is not found";
        return false;
    }

    private string? GetDocumentPath(string document)
    {
        if (string.IsNullOrWhiteSpace(document))
        {
            return null;
        }

        if (System.IO.File.Exists(document))
        {
            return document;
        }

        string webRoot = _environment.WebRootPath;
        if (string.IsNullOrWhiteSpace(webRoot))
        {
            webRoot = Path.Combine(_environment.ContentRootPath, "wwwroot");
        }

        string relative = document.TrimStart('/', '\\').Replace('/', Path.DirectorySeparatorChar);
        string candidate = Path.Combine(webRoot, relative);
        if (System.IO.File.Exists(candidate))
        {
            return candidate;
        }

        string dataCandidate = Path.Combine(webRoot, "data", Path.GetFileName(document));
        if (System.IO.File.Exists(dataCandidate))
        {
            return dataCandidate;
        }

        if (_ragService.TryResolveUploadedFile(
                Path.GetFileName(document),
                out string fullPath,
                out _,
                out _))
        {
            return fullPath;
        }

        return null;
    }
}
