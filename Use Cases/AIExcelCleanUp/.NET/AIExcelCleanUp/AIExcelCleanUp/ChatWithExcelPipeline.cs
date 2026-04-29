using System.Text;
using Azure;
using Azure.AI.OpenAI;
using OpenAI.Chat;
using Syncfusion.XlsIO;

namespace AIExcelCleaner;

/// <summary>
/// AI-Powered Excel Cleaning Pipeline
/// Workflow: Excel → CSV → AI Cleaning → Cleaned Excel
/// </summary>
public class ChatWithExcelPipeline
{
    private readonly ChatClient _chatClient;

    public ChatWithExcelPipeline()
    {
        string? apiKey = "Add your API key here";
        string? endpoint = "Add your endpoint here";
        string? modelId = "Add your model ID here";

        if (string.IsNullOrWhiteSpace(apiKey))
            throw new InvalidOperationException("Please set AZURE_OPENAI_API_KEY environment variable");
        if (string.IsNullOrWhiteSpace(endpoint))
            throw new InvalidOperationException("Please set AZURE_OPENAI_ENDPOINT environment variable");
        if (string.IsNullOrWhiteSpace(modelId))
            throw new InvalidOperationException("Please set OPENAI_MODEL environment variable");

        var azureClient = new AzureOpenAIClient(new Uri(endpoint), new AzureKeyCredential(apiKey));
        _chatClient = azureClient.GetChatClient(modelId);
    }

    private static void ValidateExcelFile(string excelFilePath)
    {
        if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
            throw new FileNotFoundException($"Excel file not found: {excelFilePath}");

        if (!excelFilePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !excelFilePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("File must be an Excel file (.xlsx or .xls)");
    }

    private static string BuildCsvContext(string excelFilePath)
    {
        using var excelEngine = new ExcelEngine();
        var excelApp = excelEngine.Excel;
        excelApp.DefaultVersion = ExcelVersion.Xlsx;

        using var fileStream = File.OpenRead(excelFilePath);
        var workbook = excelApp.Workbooks.Open(fileStream);

        var csvBuilder = new StringBuilder();

        foreach (var worksheet in workbook.Worksheets)
        {
            using var csvStream = new MemoryStream();
            worksheet.SaveAs(csvStream, ",");
            string csvData = Encoding.UTF8.GetString(csvStream.ToArray());
            csvBuilder.Append(csvData);
        }

        return csvBuilder.ToString();
    }

    private async Task<string> SendToAzureOpenAIAsync(string csvContent, string systemPrompt, string userPrompt)
    {
        var chatHistory = new List<ChatMessage>
        {
            new SystemChatMessage(systemPrompt),
            new UserChatMessage($"{userPrompt}\n\n--- EXCEL DATA (CSV FORMAT) ---\n{csvContent}")
        };

        var options = new ChatCompletionOptions
        {
            Temperature = 0.3f,
            MaxOutputTokenCount = 8000
        };

        var response = await _chatClient.CompleteChatAsync(chatHistory, options);
        return response.Value.Content[0].Text ?? string.Empty;
    }

    private static void SaveResponseToExcel(string csvResponse, string outputPath)
    {
        using var excelEngine = new ExcelEngine();
        var excelApp = excelEngine.Excel;
        excelApp.DefaultVersion = ExcelVersion.Xlsx;

        var workbook = excelApp.Workbooks.Create(1);
        var worksheet = workbook.Worksheets[0];
        worksheet.Name = "Cleaned Data";

        string csvContent = ExtractCsvFromResponse(csvResponse);

        // Split CSV into lines
        var lines = csvContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
        var validLines = lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();

        // Write directly to Excel
        int row = 1;
        foreach (var line in validLines)
        {
            var cells = line.Split(',');
            for (int col = 0; col < cells.Length; col++)
                worksheet[row, col + 1].Value = cells[col].Trim();
            row++;
        }

        // Auto-fit columns
        for (int col = 1; col <= worksheet.UsedRange.LastColumn; col++)
            worksheet.AutofitColumn(col);

        workbook.SaveAs(outputPath);
        workbook.Close();
        Console.WriteLine($"✓ Cleaned Excel saved: {outputPath}");
    }

    private static string ExtractCsvFromResponse(string aiResponse)
    {
        if (aiResponse.Contains("```csv"))
        {
            int start = aiResponse.IndexOf("```csv") + 6;
            int end = aiResponse.IndexOf("```", start);
            if (end > start) return aiResponse.Substring(start, end - start).Trim();
        }

        if (aiResponse.Contains("```"))
        {
            int start = aiResponse.IndexOf("```") + 3;
            int end = aiResponse.IndexOf("```", start);
            if (end > start) return aiResponse.Substring(start, end - start).Trim();
        }

        return aiResponse;
    }

    public async Task ExecuteChatWithExcelAsync()
    {
        Console.WriteLine("AI-Powered Excel Cleaner (Excel → CSV → AI → Cleaned Excel)\n");

        Console.Write("Enter Excel file path: ");
        string? filePath = Console.ReadLine()?.Trim().Trim('"');

        try
        {
            ValidateExcelFile(filePath ?? string.Empty);

            Console.WriteLine("Converting to CSV...");
            string csvContent = BuildCsvContext(filePath!);

            Console.WriteLine("Cleaning with AI...");
            string systemPrompt = "You are an expert Excel data cleaning assistant. Clean and standardize the provided data. Return ONLY the cleaned CSV data, nothing else.";
            string userPrompt = "Please clean this Excel data. Return ONLY cleaned CSV data.";

            string aiResponse = await SendToAzureOpenAIAsync(csvContent, systemPrompt, userPrompt);

            string outputPath = Path.Combine(
                Path.GetDirectoryName(filePath) ?? ".",
                Path.GetFileNameWithoutExtension(filePath) + "_cleaned.xlsx");

            SaveResponseToExcel(aiResponse, outputPath);
            Console.WriteLine("✓ Complete!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error: {ex.Message}");
        }
    }
}
