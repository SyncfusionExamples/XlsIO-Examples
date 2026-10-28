#nullable enable

namespace AgentRAGWithAzure.Services;

/// <summary>
/// Adds the owner column on startup and deletes every uploaded document
/// (database chunks and file) 24 hours after it was uploaded.
/// </summary>
public sealed class DocumentCleanupService : BackgroundService
{
    private static readonly TimeSpan Retention = TimeSpan.FromHours(24);
    private static readonly TimeSpan Interval = TimeSpan.FromMinutes(15);

    private readonly RagService _ragService;
    private readonly ILogger<DocumentCleanupService> _logger;

    public DocumentCleanupService(
        RagService ragService,
        ILogger<DocumentCleanupService> logger)
    {
        _ragService = ragService;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        try
        {
            await _ragService.EnsureSchemaAsync();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Could not add owner_id to rag_chunks.");
        }

        using var timer = new PeriodicTimer(Interval);
        do
        {
            try
            {
                await _ragService.DeleteExpiredDocumentsAsync(Retention);
            }
            catch (Exception ex) when (ex is not OperationCanceledException)
            {
                _logger.LogError(ex, "Expired document cleanup failed.");
            }
        }
        while (await timer.WaitForNextTickAsync(stoppingToken));
    }
}
