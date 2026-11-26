namespace AgentChatApp.Models;

/// <summary>
/// Stores token usage information for a user.
/// </summary>
public class UserTokenInfo
{
    /// <summary>
    /// Unique identifier for the user (fingerprint or authenticated user ID).
    /// </summary>
    public string UserId { get; set; } = string.Empty;

    /// <summary>
    /// Date and time when the user's tokens were last reset or first recorded.
    /// </summary>
    public DateTime DateOfLogin { get; set; }

    /// <summary>
    /// Number of tokens remaining for the current 24-hour period.
    /// </summary>
    public int RemainingTokens { get; set; }
}
