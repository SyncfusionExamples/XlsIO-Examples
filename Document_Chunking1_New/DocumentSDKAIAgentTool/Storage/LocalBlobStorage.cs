
using Syncfusion.AI.AgentTools.Core;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AgentChatApp.Storage;

/// <summary>
/// Local file system implementation of <see cref="IDocumentStorage"/>.
/// Read/load operations use the <c>Input</c> subfolder;
/// write/upload operations use the <c>Output</c> subfolder.
/// </summary>
public sealed class LocalBlobStorage : IDocumentStorage
{
    private readonly string _inputFolder;
    private readonly string _outputFolder;

    public LocalBlobStorage(string dataFolder)
    {
        ArgumentException.ThrowIfNullOrEmpty(dataFolder);
        _inputFolder  = Path.Combine(dataFolder, "Input");
        _outputFolder = Path.Combine(dataFolder, "Output");
        Directory.CreateDirectory(_inputFolder);
        Directory.CreateDirectory(_outputFolder);
    }

    /// <inheritdoc/>
    public Stream Read(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        filePath = filePath.Replace('\\', '/');
        // Determine the base folder from the path (Input/ or Output/)
        string fullPath;
        if (filePath.StartsWith("Input/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Input/".Length);
            fullPath = Path.Combine(_inputFolder, relativePath);
        }
        else if (filePath.StartsWith("Output/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Output/".Length);
            fullPath = Path.Combine(_outputFolder, relativePath);
        }
        else
        {
            // Try both folders if no prefix
            var inputPath = Path.Combine(_inputFolder, filePath);
            var outputPath = Path.Combine(_outputFolder, filePath);
            if (File.Exists(inputPath))
                fullPath = inputPath;
            else if (File.Exists(outputPath))
                fullPath = outputPath;
            else
                throw new FileNotFoundException($"File not found: {filePath}");
        }
        
        if (!File.Exists(fullPath))
            throw new FileNotFoundException($"File not found: {filePath}");
            
        var ms = new MemoryStream(File.ReadAllBytes(fullPath));
        ms.Position = 0;
        return ms;
    }

    /// <inheritdoc/>
    public bool Write(string filePath, Stream documentStream)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        ArgumentNullException.ThrowIfNull(documentStream);
        filePath = filePath.Replace('\\', '/');
        try
        {
            documentStream.Position = 0;
            
            // Determine the base folder from the path (Input/ or Output/)
            string fullPath;
            if (filePath.StartsWith("Input/", StringComparison.OrdinalIgnoreCase))
            {
                var relativePath = filePath.Substring("Input/".Length);
                fullPath = Path.Combine(_inputFolder, relativePath);
            }
            else if (filePath.StartsWith("Output/", StringComparison.OrdinalIgnoreCase))
            {
                var relativePath = filePath.Substring("Output/".Length);
                fullPath = Path.Combine(_outputFolder, relativePath);
            }
            else
            {
                // Default to Output folder if no prefix
                fullPath = Path.Combine(_outputFolder, filePath);
            }
            
            var dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(dir))
                Directory.CreateDirectory(dir);
            using (var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
            {
                documentStream.CopyTo(fs);
            }
            return true;
        }
        catch
        {
            return false;
        }
    }



    /// <inheritdoc/>
    public bool Exists(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        filePath = filePath.Replace('\\', '/');
        // Determine the base folder from the path (Input/ or Output/)
        if (filePath.StartsWith("Input/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Input/".Length);
            return File.Exists(Path.Combine(_inputFolder, relativePath));
        }
        else if (filePath.StartsWith("Output/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Output/".Length);
            return File.Exists(Path.Combine(_outputFolder, relativePath));
        }
        else
        {
            // Check both Input and Output folders if no prefix
            return File.Exists(Path.Combine(_inputFolder, filePath))
                || File.Exists(Path.Combine(_outputFolder, filePath));
        }
    }

    /// <summary>
    /// Returns metadata for all files in both Input and Output folders.
    /// Used by the Files API to populate the Documents folder listing.
    /// </summary>
    public IReadOnlyList<BlobItemInfo> GetAllBlobItems()
    {
        var inputFiles  = Directory.GetFiles(_inputFolder,  "*", SearchOption.AllDirectories)
            .Select(f => (file: f, rel: "Input/"  + Path.GetRelativePath(_inputFolder,  f).Replace("\\", "/")));
        var outputFiles = Directory.GetFiles(_outputFolder, "*", SearchOption.AllDirectories)
            .Select(f => (file: f, rel: "Output/" + Path.GetRelativePath(_outputFolder, f).Replace("\\", "/")));

        return inputFiles.Concat(outputFiles)
            .Select(x => new BlobItemInfo(x.rel, new FileInfo(x.file).Length, GetContentType(x.file), File.GetLastWriteTime(x.file)))
            .OrderBy(b => b.Name)
            .ToList();
    }

    /// <summary>
    /// Returns a hierarchical structure with Input and Output as top-level folders.
    /// </summary>
    public FileSystemStructure GetFileSystemStructure()
    {
        var root = new FileSystemStructure();

        foreach (var (folder, label) in new[] { (_inputFolder, "Input"), (_outputFolder, "Output") })
        {
            var folderInfo = new FolderInfo { Name = label, Files = new List<BlobItemInfo>() };
            foreach (var file in Directory.GetFiles(folder, "*", SearchOption.AllDirectories))
            {
                var relPath = label + "/" + Path.GetRelativePath(folder, file).Replace("\\", "/");
                folderInfo.Files.Add(new BlobItemInfo(
                    relPath,
                    new FileInfo(file).Length,
                    GetContentType(file),
                    File.GetLastWriteTime(file)));
            }
            root.Folders[label] = folderInfo;
        }

        return root;
    }

    /// <summary>
    /// Uploads a stream as a new file. Determines the target folder from the filePath.
    /// Used by the Files API upload endpoint.
    /// </summary>
    public void Upload(string filePath, Stream stream, string contentType)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        ArgumentNullException.ThrowIfNull(stream);
        stream.Position = 0;
        filePath = filePath.Replace('\\', '/');
        // Determine the base folder from the path (Input/ or Output/)
        string fullPath;
        if (filePath.StartsWith("Input/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Input/".Length);
            fullPath = Path.Combine(_inputFolder, relativePath);
        }
        else if (filePath.StartsWith("Output/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Output/".Length);
            fullPath = Path.Combine(_outputFolder, relativePath);
        }
        else
        {
            // Default to Input folder if no prefix
            fullPath = Path.Combine(_inputFolder, filePath);
        }
        
        var dir = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        using (var fs = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
        {
            stream.CopyTo(fs);
        }
    }

    /// <summary>
    /// Downloads a file as a byte array. Determines the folder from the filePath prefix.
    /// </summary>
    public byte[] Download(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        filePath = filePath.Replace('\\', '/');
        // Determine the base folder from the path (Input/ or Output/)
        string fullPath;
        if (filePath.StartsWith("Input/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Input/".Length);
            fullPath = Path.Combine(_inputFolder, relativePath);
        }
        else if (filePath.StartsWith("Output/", StringComparison.OrdinalIgnoreCase))
        {
            var relativePath = filePath.Substring("Output/".Length);
            fullPath = Path.Combine(_outputFolder, relativePath);
        }
        else
        {
            // Try both folders if no prefix
            var inputPath = Path.Combine(_inputFolder, filePath);
            var outputPath = Path.Combine(_outputFolder, filePath);
            if (File.Exists(inputPath))  return File.ReadAllBytes(inputPath);
            if (File.Exists(outputPath)) return File.ReadAllBytes(outputPath);
            throw new FileNotFoundException($"File not found: {filePath}");
        }
        
        if (File.Exists(fullPath))
            return File.ReadAllBytes(fullPath);
            
        throw new FileNotFoundException($"File not found: {filePath}");
    }
    
    /// <inheritdoc/>
    public bool Delete(string filePath)
    {
        ArgumentException.ThrowIfNullOrEmpty(filePath);
        filePath = filePath.Replace('\\', '/');
        try
        {
            // Determine the base folder from the path (Input/ or Output/)
            string fullPath;
            if (filePath.StartsWith("Input/", StringComparison.OrdinalIgnoreCase))
            {
                var relativePath = filePath.Substring("Input/".Length);
                fullPath = Path.Combine(_inputFolder, relativePath);
            }
            else if (filePath.StartsWith("Output/", StringComparison.OrdinalIgnoreCase))
            {
                var relativePath = filePath.Substring("Output/".Length);
                fullPath = Path.Combine(_outputFolder, relativePath);
            }
            else
            {
                // Try both folders if no prefix
                foreach (var folder in new[] { _outputFolder, _inputFolder })
                {
                    var testPath = Path.Combine(folder, filePath);
                    if (File.Exists(testPath))
                    {
                        File.Delete(testPath);
                        return true;
                    }
                }
                return false;
            }
            
            if (File.Exists(fullPath))
            {
                File.Delete(fullPath);
                return true;
            }
            return false;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Failed to delete file '{filePath}': {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Gets the content type based on file extension.
    /// </summary>
    private static string GetContentType(string filePath)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        return ext switch
        {
            ".txt" => "text/plain",
            ".json" => "application/json",
            ".md" => "text/markdown",
            ".pdf" => "application/pdf",
            ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".html" => "text/html",
            ".htm" => "text/html",
            ".jpg" => "image/jpeg",
            ".jpeg" => "image/jpeg",
            ".png" => "image/png",
            ".gif" => "image/gif",
            _ => "application/octet-stream"
        };
    }
}

/// <summary>Lightweight blob metadata DTO returned by <see cref="LocalBlobStorage.GetAllBlobItems"/>.</summary>
public sealed record BlobItemInfo(string Name, long Size, string ContentType, DateTime LastModified);

/// <summary>Represents the file system structure with folders and files.</summary>
public sealed class FileSystemStructure
{
    public Dictionary<string, FolderInfo> Folders { get; set; } = new();
    public List<BlobItemInfo> Files { get; set; } = new();
}

/// <summary>Represents a folder with its contained files.</summary>
public sealed class FolderInfo
{
    public string Name { get; set; } = string.Empty;
    public List<BlobItemInfo> Files { get; set; } = new();
}

