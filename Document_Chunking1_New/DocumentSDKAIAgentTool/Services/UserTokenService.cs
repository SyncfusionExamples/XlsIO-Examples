using AgentChatApp.Models;
using System.Text.Json;

namespace AgentChatApp.Services;

/// <summary>
/// Manages user token limits for AI API usage.
/// Tracks token consumption per user and enforces daily limits.
/// </summary>
public class UserTokenService
{
    private const string TokenFilePath = "user_tokens.json";
    
    // HARDCODED FOR TESTING - Change before publishing!
    private const int _dailyTokenLimit = 300;          // 100 tokens for testing
    private const double _resetHours = 24;             // 24 hours reset for testing

    // Use Local timezone so reset times display correctly for users
    private static readonly TimeZoneInfo IndianStandardTime = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
    private static readonly TimeZoneInfo ApplicationTimeZone = IndianStandardTime;
    private readonly SemaphoreSlim _fileLock = new(1, 1);

    public UserTokenService(IConfiguration configuration)
    {
        // Configuration values ignored - using hardcoded testing values above
    }

    /// <summary>
    /// Initializes a new user in the token system if they don't already exist.
    /// </summary>
    public async Task InitializeUserAsync(string userCode)
    {
        await _fileLock.WaitAsync();
        try
        {
            var tokenData = await ReadTokensFromFileAsync();

            if (!tokenData.ContainsKey(userCode))
            {
                tokenData[userCode] = new UserTokenInfo
                {
                    UserId = userCode,
                    DateOfLogin = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, ApplicationTimeZone),
                    RemainingTokens = _dailyTokenLimit
                };
                await WriteTokensToFileAsync(tokenData);
            }
        }
        finally
        {
            _fileLock.Release();
        }
    }

    /// <summary>
    /// Gets remaining tokens for a user, resets if the configured period has passed.
    /// </summary>
    public async Task<int> GetRemainingTokensAsync(string userCode)
    {
        var tokens = await CheckAndResetTokensAsync(userCode);
        return tokens.ContainsKey(userCode)
            ? tokens[userCode].RemainingTokens
            : _dailyTokenLimit;
    }

    /// <summary>
    /// Updates token count after AI API usage.
    /// </summary>
    public async Task UpdateTokensAsync(string userCode, int remainingTokens)
    {
        await _fileLock.WaitAsync();
        try
        {
            var tokenData = await ReadTokensFromFileAsync();

            if (tokenData.ContainsKey(userCode))
            {
                tokenData[userCode].RemainingTokens = remainingTokens;
            }
            else
            {
                tokenData[userCode] = new UserTokenInfo
                {
                    UserId = userCode,
                    DateOfLogin = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, ApplicationTimeZone),
                    RemainingTokens = remainingTokens
                };
            }

            await WriteTokensToFileAsync(tokenData);
        }
        finally
        {
            _fileLock.Release();
        }
    }

    /// <summary>
    /// Checks if the reset period has passed and resets tokens if needed.
    /// </summary>
    public async Task<Dictionary<string, UserTokenInfo>> CheckAndResetTokensAsync(string userCode)
    {
        await _fileLock.WaitAsync();
        try
        {
            var tokenData = await ReadTokensFromFileAsync();

            if (tokenData.ContainsKey(userCode))
            {
                var userTokenInfo = tokenData[userCode];
                var currentTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, ApplicationTimeZone);
                var timeDifference = currentTime - userTokenInfo.DateOfLogin;

                // Using hours for testing instead of minutes
                if (timeDifference.TotalHours >= _resetHours)
                {
                    userTokenInfo.RemainingTokens = _dailyTokenLimit;
                    userTokenInfo.DateOfLogin = currentTime;
                    await WriteTokensToFileAsync(tokenData);
                }
            }

            return tokenData;
        }
        finally
        {
            _fileLock.Release();
        }
    }

    /// <summary>
    /// Generates an alert message with reset time information.
    /// </summary>
    public async Task<string> GetAlertMessageAsync(string userCode)
    {
        await _fileLock.WaitAsync();
        try
        {
            var tokenData = await ReadTokensFromFileAsync();

            if (tokenData.ContainsKey(userCode))
            {
                var userTokenInfo = tokenData[userCode];
                // Calculate reset time (using configured _resetHours for testing)
                var resetDateTime = userTokenInfo.DateOfLogin.AddHours(_resetHours);
                // Format: "Friday, June 5, 2026 2:30 PM" (in local timezone)
                var resetTimeFormatted = resetDateTime.ToString("dddd, MMMM d, yyyy h:mm tt");
                
                var message = $"You have reached your token limit. Your tokens will reset on {resetTimeFormatted}. " +
                    $"Download our <a href=\"https://github.com/syncfusion/document-sdk-ai-agent-tools/tree/master/Examples/Showcase\" target=\"_blank\" style=\"color: #721c24; text-decoration: underline;\">Syncfusion Document SDK AI Agent Tools</a> from GitHub to explore this sample locally with your own API key.";
                return message;
            }

            // Return default message if user not found in token data
            var defaultResetTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, ApplicationTimeZone).AddHours(_resetHours);
            var defaultResetFormatted = defaultResetTime.ToString("dddd, MMMM d, yyyy h:mm tt");
            return $"You have reached your token limit. Your tokens will reset on {defaultResetFormatted}. " +
                   $"Download our <a href=\"https://github.com/syncfusion/document-sdk-ai-agent-tools/tree/master/Examples\" target=\"_blank\" style=\"color: #721c24; text-decoration: underline;\">Syncfusion Document SDK AI Agent Tools</a> from GitHub to explore this sample locally with your own API key.";
        }
        finally
        {
            _fileLock.Release();
        }
    }

    /// <summary>
    /// Reads token data from JSON file.
    /// </summary>
    private async Task<Dictionary<string, UserTokenInfo>> ReadTokensFromFileAsync()
    {
        if (!File.Exists(TokenFilePath))
        {
            var initialData = new Dictionary<string, UserTokenInfo>();
            await WriteTokensToFileAsync(initialData);
            return initialData;
        }

        try
        {
            var json = await File.ReadAllTextAsync(TokenFilePath);
            var tokenData = JsonSerializer.Deserialize<Dictionary<string, UserTokenInfo>>(json);
            return tokenData ?? new Dictionary<string, UserTokenInfo>();
        }
        catch (JsonException)
        {
            // If file is corrupted, start fresh
            return new Dictionary<string, UserTokenInfo>();
        }
    }

    /// <summary>
    /// Writes token data to JSON file.
    /// </summary>
    private async Task WriteTokensToFileAsync(Dictionary<string, UserTokenInfo> tokenData)
    {
        var json = JsonSerializer.Serialize(tokenData, new JsonSerializerOptions { WriteIndented = true });
        await File.WriteAllTextAsync(TokenFilePath, json);
    }
}
