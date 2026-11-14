#nullable enable

using System.Text.Json;
using System.Text.RegularExpressions;

namespace AgentRAGWithAzure.Services;

/// <summary>
/// Limits each user to one document upload per 24 hours.
/// The allowance resets 24 hours after the last successful upload.
/// </summary>
public sealed partial class UploadQuotaService
{
    private static readonly TimeSpan ResetPeriod = TimeSpan.FromHours(24);

    private readonly string _filePath;
    private readonly SemaphoreSlim _lock = new(1, 1);

    public UploadQuotaService(IWebHostEnvironment environment)
    {
        string folder = Path.Combine(environment.ContentRootPath, "App_Data");
        Directory.CreateDirectory(folder);
        _filePath = Path.Combine(folder, "upload_quota.json");
    }

    [GeneratedRegex("^[A-Za-z0-9_-]{4,64}$")]
    private static partial Regex UserCodePattern();

    public static bool IsValidUserCode(string? userCode) =>
        !string.IsNullOrEmpty(userCode) && UserCodePattern().IsMatch(userCode);

    /// <summary>Returns null when the user may upload, otherwise the reset time (UTC).</summary>
    public async Task<DateTime?> GetBlockedUntilAsync(string userCode)
    {
        await _lock.WaitAsync();
        try
        {
            var data = await ReadAsync();
            if (data.TryGetValue(userCode, out DateTime lastUpload))
            {
                DateTime resetAt = lastUpload + ResetPeriod;
                if (DateTime.UtcNow < resetAt)
                {
                    return resetAt;
                }
            }
            return null;
        }
        finally
        {
            _lock.Release();
        }
    }

    public async Task RecordUploadAsync(string userCode)
    {
        await _lock.WaitAsync();
        try
        {
            var data = await ReadAsync();
            data[userCode] = DateTime.UtcNow;
            await File.WriteAllTextAsync(_filePath, JsonSerializer.Serialize(data));
        }
        finally
        {
            _lock.Release();
        }
    }

    private async Task<Dictionary<string, DateTime>> ReadAsync()
    {
        if (!File.Exists(_filePath))
        {
            return new Dictionary<string, DateTime>();
        }

        try
        {
            string json = await File.ReadAllTextAsync(_filePath);
            return JsonSerializer.Deserialize<Dictionary<string, DateTime>>(json)
                   ?? new Dictionary<string, DateTime>();
        }
        catch (JsonException)
        {
            return new Dictionary<string, DateTime>();
        }
    }
}
