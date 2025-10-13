using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using OpenAI;
using OpenAI.Chat;
using System.ClientModel;

/// <summary>
/// HumanExcelChatBot: An AI-powered chatbot for summarizing and querying Excel workbooks using Syncfusion XlsIO and OpenAI.
/// </summary>
class HumanExcelChatBot
{
    /// <summary>
    /// Entry point for the HumanExcelChatBot application.
    /// </summary>
    static async Task Main()
    {
        // Replace with your actual OpenAI API key or set it in environment variables for security
        string? openAIApiKey = "OpenAI API Key";
        await ExecuteChatWithExcel(openAIApiKey);
    }

    /// <summary>
    /// Execute chat with Excel document.
    /// </summary>
    private async static Task ExecuteChatWithExcel(string openAIApiKey)
    {
        Console.WriteLine("AI Powered Excel ChatBot");

        Console.WriteLine("Enter full Excel file path (e.g., C:\\Data\\report.xlsx):");

        //Read user input for Excel file path
        string? excelFilePath = Console.ReadLine()?.Trim().Trim('"');

        if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
        {
            Console.WriteLine("Invalid path. Exiting.");
            return;
        }


        if (string.IsNullOrWhiteSpace(openAIApiKey))
        {
            Console.WriteLine("OPENAI_API_KEY not set. Exiting.");
            return;
        }

        string csvContext;
        try
        {
            // Convert Excel to CSV-like context
            csvContext = BuildCsvContext(excelFilePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read Excel: {ex.Message}");
            return;
        }

        Console.WriteLine("\nOverview (AI):");

        List<ChatMessage> chatHistory = new List<ChatMessage>
        {
            new SystemChatMessage("You are an assistant that helps analyze Excel data.")
        };

        try
        {
            // Get an overview of the Excel data
            string overview = await AskOpenAIAsync(
                openAIApiKey,
                model: "gpt-5", // replace with a model available in your account
                systemPrompt: "You are a helpful assistant. Provide a concise 3-6 bullet summary of the workbook data.",
                userContent: csvContext, chatHistory);
            Console.WriteLine(overview);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"OpenAI overview failed: {ex.Message}");
        }

        Console.WriteLine("\nDo you want any specific details? Type 'Stop' to exit.");



        while (true)
        {
            // Read user question
            Console.Write("\nYou: ");

            string? userQuestion = Console.ReadLine();

            if (userQuestion == null)
                continue;

            if (userQuestion.Trim().Equals("Stop", StringComparison.OrdinalIgnoreCase)) 
                break;

            // Define system prompt for the chatbot
            string systemPrompt = @"You are an intelligent Excel assistant built using Syncfusion XlsIO and OpenAI. Your job is to help users understand, analyze, and interact with Excel data through natural language. You are integrated into a C# application and can access structured Excel content such as tables, charts, formulas, and cell values.
                Your capabilities include:
                - Summarizing Excel sheets and ranges
                - Answering questions about data trends, totals, averages, and comparisons
                - Explaining formulas and calculations
                - Suggesting improvements or insights based on the data
                - Responding in a friendly, professional, and concise manner
                Constraints:
                You do not generate or modify Excel files directly. Instead, you interpret and explain the data provided to you by the application.
                Always assume the user is referring to the most recent Excel content unless stated otherwise. Maintain context across the conversation to provide coherent and helpful responses.
                If the user asks something outside the scope of Excel data, politely redirect them back to Excel-related tasks.";

            string userPrompt =
                "Question:\n" + userQuestion;

            try
            {
                // Get answer from OpenAI based on user question and chat history
                string answer = await AskOpenAIAsync(openAIApiKey, "gpt-5", systemPrompt, userPrompt, chatHistory);
                Console.WriteLine("ChatBot: " + answer);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"OpenAI error: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Builds a CSV-like text context from the specified Excel file, limited to a maximum number of rows per sheet.
    /// </summary>
    /// <param name="maxRowsPerSheet">Maximum number of rows to include per sheet.</param>
    /// <returns>CSV-like string representation of the workbook.</returns>
    private static string BuildCsvContext(string excelFilePath)
    {
        // Initialize Syncfusion Excel engine
        using ExcelEngine excelEngine = new ExcelEngine();

        // Set default version to XLSX
        IApplication excelApp = excelEngine.Excel;
        excelApp.DefaultVersion = ExcelVersion.Xlsx;

        // Open the Excel file
        using FileStream fileStream = File.OpenRead(excelFilePath);
        IWorkbook workbook = excelApp.Workbooks.Open(fileStream);

        // Build CSV-like context
        StringBuilder stringBuilder = new StringBuilder();

        // Add file and sheet info
        stringBuilder.AppendLine($"File: {excelFilePath}");
        stringBuilder.AppendLine($"Sheets: {workbook.Worksheets.Count}");
        stringBuilder.AppendLine();

        // Convert each worksheet to CSV format and append to the context
        foreach (IWorksheet worksheet in workbook.Worksheets)
        {
            MemoryStream csvStream = new MemoryStream();

            // Save workbook as CSV
            worksheet.SaveAs(csvStream, ",");

            // Convert CSV to text
            string excelData = Encoding.UTF8.GetString(csvStream.ToArray());
            stringBuilder.AppendLine("Sheet Name : " + worksheet.Name);
            stringBuilder.AppendLine();
            stringBuilder.AppendLine(excelData);
            stringBuilder.AppendLine();
        }

        return stringBuilder.ToString();
    }

    /// <summary>
    /// Sends a chat completion request to OpenAI and returns the response.
    /// </summary>
    /// <param name="apiKey">OpenAI API key.</param>
    /// <param name="model">Model name.</param>
    /// <param name="systemPrompt">System prompt for the assistant.</param>
    /// <param name="userContent">User content/question.</param>
    /// <param name="chatHistory">Chat History</param>
    /// <returns>AI-generated response as a string.</returns>
    private static async Task<string> AskOpenAIAsync(string apiKey, string model, string systemPrompt, string userContent, List<ChatMessage> chatHistory)
    {
        // Initialize OpenAI client
        OpenAIClient openAIClient = new OpenAIClient(apiKey);

        // Create chat client for the specified model
        ChatClient chatClient = openAIClient.GetChatClient(model);

        // Append system and user messages to chat history
        chatHistory.Add(new SystemChatMessage(systemPrompt));

        // Add user message to chat history
        chatHistory.Add(new UserChatMessage(userContent));

        //Get AI response
        ClientResult<ChatCompletion> chatResult = await chatClient.CompleteChatAsync(chatHistory);

        string response = chatResult.Value.Content[0].Text ?? string.Empty;

        // Add assistant response to chat history
        chatHistory.Add(new AssistantChatMessage(response));

        return response;
    }
}