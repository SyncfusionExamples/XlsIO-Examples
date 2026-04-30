using AIExcelCleaner;

try
{
    ChatWithExcelPipeline pipeline = new ChatWithExcelPipeline();
    await pipeline.ExecuteChatWithExcelAsync();
}
catch (InvalidOperationException ex) when (ex.Message.Contains("Azure"))
{
    Console.WriteLine($"\n Configuration Error: {ex.Message}\n");
    Console.WriteLine("Please set Azure OpenAI credentials:");
    Console.WriteLine("  $env:AZURE_OPENAI_API_KEY = \"your-api-key\"");
    Console.WriteLine("  $env:AZURE_OPENAI_ENDPOINT = \"your-endpoint\"");
    Console.WriteLine("  $env:OPENAI_MODEL = \"your-model-id\"\n");
}
catch (Exception ex)
{
    Console.WriteLine($"\n Error: {ex.Message}");
    if (ex.InnerException != null)
        Console.WriteLine($"Details: {ex.InnerException.Message}");
}
