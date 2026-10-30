using AgentChatApp.Services;
using AgentChatApp.Storage;
using Microsoft.AspNetCore.Mvc;

namespace AgentChatApp.Controllers;

/// <summary>
/// Serves the File Explorer panel. Lists, downloads, uploads, and deletes blobs
/// in Azure Blob Storage. The UI exposes folders (Input, Output, etc.) that map
/// to folder prefixes in the configured blob container.
/// </summary>
[ApiController]
[Route("api/[controller]")]
public class FilesController : ControllerBase
{
    private readonly LocalBlobStorage _blobStorage;

    public FilesController(AgentService agentService)
    {
        _blobStorage = agentService.BlobStorage;
    }

    /// <summary>Returns the list of folder names for the File Explorer sidebar.</summary>
    [HttpGet("folders")]
    public IActionResult GetFolders()
    {
        try
        {
            var structure = _blobStorage.GetFileSystemStructure();
            var folders = new List<object>();
            
            // Return only the subfolders (Input, Output) - no Documents folder
            foreach (var folder in structure.Folders.Keys.OrderBy(f => f))
            {
                folders.Add(new { name = folder });
            }
            
            return Ok(folders);
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>Returns all blobs in a specific folder (Input, Output, etc.).</summary>
    [HttpGet("folders/{folderName}")]
    public IActionResult GetFiles(string folderName)
    {
        try
        {
            var structure = _blobStorage.GetFileSystemStructure();
            
            // Return specific folder contents
            if (!structure.Folders.ContainsKey(folderName))
                return NotFound(new { error = $"Folder '{folderName}' not found." });

            var folderFiles = structure.Folders[folderName].Files
                .Where(b => !Path.GetFileName(b.Name).Equals(".gitkeep", StringComparison.OrdinalIgnoreCase))
                .Select(b => new
                {
                    name = Path.GetFileName(b.Name),
                    fullPath = b.Name,
                    isFolder = false,
                    size = FormatFileSize(b.Size),
                    extension = Path.GetExtension(b.Name).TrimStart('.').ToUpperInvariant(),
                    modified = b.LastModified.ToString("yyyy-MM-dd")
                }).ToList();

            return Ok(new { folder = folderName, files = folderFiles });
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>Downloads a blob by name or path.</summary>
    [HttpGet("download/{folderName}/{*fileName}")]
    public IActionResult DownloadFile(string folderName, string fileName)
    {
        var filePath = $"{folderName}/{Uri.UnescapeDataString(fileName)}";

        if (!_blobStorage.Exists(filePath))
            return NotFound(new { error = $"Document '{filePath}' not found." });

        try
        {
            var bytes = _blobStorage.Download(filePath);
            var contentType = GetContentType(filePath);
            return File(bytes, contentType, Path.GetFileName(filePath));
        }
        catch (Exception ex)
        {
            return StatusCode(500, new { error = ex.Message });
        }
    }

    /// <summary>Deletes a blob by name or path.</summary>
    [HttpDelete("delete/{folderName}/{*fileName}")]
    public IActionResult DeleteFile(string folderName, string fileName)
    {
        var filePath = $"{folderName}/{Uri.UnescapeDataString(fileName)}";

        if (!_blobStorage.Exists(filePath))
            return NotFound(new { error = $"Document '{filePath}' not found." });

        bool deleted = _blobStorage.Delete(filePath);
        if (!deleted)
            return StatusCode(500, new { error = $"Failed to delete document '{filePath}'." });

        return Ok(new { message = $"Document '{filePath}' deleted successfully." });
    }

    /// <summary>Uploads one or more files into a specified folder.</summary>
    [HttpPost("upload")]
    public async Task<IActionResult> UploadFiles(IList<IFormFile> files)
    {
        if (files == null || files.Count == 0)
            return BadRequest(new { error = "No files provided." });

        string folder = "Input";

        var uploaded = new List<string>();
        foreach (var formFile in files)
        {
            if (formFile.Length == 0) continue;

            var fileName = Path.GetFileName(formFile.FileName);
            var filePath = $"{folder}/{fileName}";
            var contentType = formFile.ContentType.NullIfEmpty() ?? GetContentType(fileName);

            await using var stream = formFile.OpenReadStream();
            using var ms = new MemoryStream();
            await stream.CopyToAsync(ms);
            ms.Position = 0;

            _blobStorage.Upload(filePath, ms, contentType);
            uploaded.Add(filePath);
        }

        return Ok(new { message = $"Uploaded {uploaded.Count} file(s) to {folder}.", files = uploaded });
    }

    // ── Helpers ────────────────────────────────────────────────────────────

    private static string GetContentType(string fileName) =>
        Path.GetExtension(fileName).ToLowerInvariant() switch
        {
            ".pdf"  => "application/pdf",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".doc"  => "application/msword",
            ".xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".xls"  => "application/vnd.ms-excel",
            ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            ".json" => "application/json",
            ".md"   => "text/markdown",
            ".txt"  => "text/plain",
            ".rtf"  => "application/rtf",
            ".html" => "application/octet-stream",
            ".xlsm" => "application/vnd.ms-excel.sheet.macroenabled.12",
            ".csv"  => "text/csv",
            ".png"  => "image/png",
            ".jpg"  => "image/jpeg",
            ".jpeg" => "image/jpeg",
            _       => "application/octet-stream"
        };

    private static string FormatFileSize(long bytes) => bytes switch
    {
        < 1024        => $"{bytes} B",
        < 1_048_576   => $"{bytes / 1024.0:F1} KB",
        _             => $"{bytes / 1_048_576.0:F1} MB"
    };
}

internal static class StringExtensions
{
    public static string? NullIfEmpty(this string? value) =>
        string.IsNullOrEmpty(value) ? null : value;
}
