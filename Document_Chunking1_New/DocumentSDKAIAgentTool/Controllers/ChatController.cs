using AgentChatApp.Services;
using Microsoft.AspNetCore.Mvc;

namespace AgentChatApp.Controllers;

[ApiController]
[Route("api/[controller]")]
public class ChatController : ControllerBase
{
    private readonly AgentService _agentService;

    public ChatController(AgentService agentService)
    {
        _agentService = agentService;
    }

    /// <summary>
    /// Sends a message to the AI agent and streams the response as plain text.
    /// Includes token limit checking before processing the request.
    /// </summary>
    [HttpPost("send")]
    public async Task SendMessage([FromBody] ChatRequest request, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(request.Message))
        {
            Response.StatusCode = 400;
            await Response.WriteAsync("Message cannot be empty.", cancellationToken);
            return;
        }

        // Check token limit if userCode is provided
        if (!string.IsNullOrWhiteSpace(request.UserCode))
        {
            var remainingTokens = await _agentService.CheckTokenLimitAsync(request.UserCode, request.Message);
            if (remainingTokens == null)
            {
                Response.StatusCode = 429; // Too Many Requests
                Response.ContentType = "application/json";
                var alertMessage = await _agentService.GetTokenLimitMessageAsync(request.UserCode);
                await Response.WriteAsync(System.Text.Json.JsonSerializer.Serialize(new
                {
                    error = "Token limit reached",
                    message = alertMessage
                }), cancellationToken);
                return;
            }
        }

        Response.ContentType = "text/plain; charset=utf-8";
        Response.Headers["Cache-Control"] = "no-cache";
        Response.Headers["X-Accel-Buffering"] = "no";

        try
        {
            await _agentService.StreamResponseAsync(
                request.SessionId,
                request.Message,
                async chunk =>
                {
                    await Response.WriteAsync(chunk, cancellationToken);
                    await Response.Body.FlushAsync(cancellationToken);
                },
                cancellationToken,
                request.UserCode);
        }
        catch (Exception ex) when (!Response.HasStarted)
        {
            Response.StatusCode = 500;
            await Response.WriteAsync($"Agent error: {ex.Message}", cancellationToken);
        }
        catch (Exception ex)
        {
            // Response already started (streaming began) — append error as a final chunk
            await Response.WriteAsync($"\n\n❌ Agent error: {ex.Message}", cancellationToken);
        }
    }

    /// <summary>
    /// Gets the remaining token count for a user.
    /// Initializes the user if they don't already exist (called on page load).
    /// </summary>
    [HttpGet("tokens/{userCode}")]
    public async Task<IActionResult> GetTokenInfo(string userCode)
    {
        // Initialize user in the system when they first visit the page
        await _agentService.InitializeUserAsync(userCode);
        
        var remainingTokens = await _agentService.CheckTokenLimitAsync(userCode, "");
        if (remainingTokens == null)
        {
            var alertMessage = await _agentService.GetTokenLimitMessageAsync(userCode);
            return Ok(new
            {
                remainingTokens = 0,
                limitReached = true,
                message = alertMessage
            });
        }

        return Ok(new
        {
            remainingTokens = remainingTokens.Value,
            limitReached = false,
            message = ""
        });
    }

    /// <summary>
    /// Clears the conversation history for the given session.
    /// </summary>
    [HttpDelete("session/{sessionId}")]
    public async Task<IActionResult> ClearSession(string sessionId, CancellationToken cancellationToken)
    {
        await _agentService.ClearSessionAsync(sessionId, cancellationToken);
        return Ok(new { cleared = true });
    }
}

public sealed record ChatRequest(string SessionId, string Message, string? UserCode = null);
