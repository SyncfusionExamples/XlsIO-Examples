using Syncfusion.XlsIO;
using Azure.AI.OpenAI;
using OpenAI.Chat;
using System.ClientModel;
using System.Text;

namespace ExcelChatBot
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            // Set the Azure OpenAI API key
            string azureOpenAIApiKey = "Replace your Azure OpenAI key";

            // Start the Excel chatbot workflow
            await ExecuteChatWithExcel(azureOpenAIApiKey);
        }

        // Performs AI-powered interactions with Excel data
        private async static Task ExecuteChatWithExcel(string azureOpenAIApiKey)
        {
            // Display the application title
            Console.WriteLine("AI-Powered Excel ChatBot");

            // Prompt the user for the Excel file path
            Console.WriteLine("Enter full Excel file path (e.g., C:\\Data\\report.xlsx):");

            // Read and clean the user input
            string? excelFilePath = Console.ReadLine()?.Trim().Trim('"');

            // Validate the Excel file path
            if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
            {
                // Notify the user that the path is invalid
                Console.WriteLine("Invalid path. Exiting.");

                // Stop execution
                return;
            }

            // Validate the Azure OpenAI API key
            if (string.IsNullOrWhiteSpace(azureOpenAIApiKey))
            {
                // Notify the user that the API key is missing
                Console.WriteLine("AZURE_OPENAI_API_KEY not set. Exiting.");

                // Stop execution
                return;
            }

            // Stores workbook data converted into text
            string csvContext;

            try
            {
                // Extract workbook content as CSV text
                csvContext = BuildCsvContext(excelFilePath);
            }
            catch (Exception ex)
            {
                // Display the workbook reading error
                Console.WriteLine($"Failed to read Excel: {ex.Message}");

                // Stop execution
                return;
            }

            // Display the overview heading
            Console.WriteLine("\nOverview (AI):");

            // Create chat history to maintain conversation context
            List<ChatMessage> chatHistory = new List<ChatMessage>
            {
                // Define the assistant's role
                new SystemChatMessage("You are an assistant that helps analyze Excel data.")
            };

            try
            {
                // Generate a workbook summary using Azure OpenAI
                string overview = await AskAzureOpenAIAsync(
                    azureOpenAIApiKey,
                    model: "your-model-name",
                    systemPrompt: "You are a helpful assistant. Provide a concise 3-6 bullet summary of the workbook data.",
                    userContent: csvContext, chatHistory
                    );

                // Display the generated summary
                Console.WriteLine(overview);
            }
            catch (Exception ex)
            {
                // Display summary generation errors
                Console.WriteLine($"Azure OpenAI overview failed: {ex.Message}");
            }

            // Prompt the user for follow-up questions
            Console.WriteLine("\nDo you want any specific details? Type 'Stop' to exit.");

            // Continue receiving user questions
            while (true)
            {
                // Show the user input prompt
                Console.Write("\nYou: ");

                // Read the user question
                string? userQuestion = Console.ReadLine();

                // Ignore empty input
                if (userQuestion == null)
                    continue;

                // Exit the loop if the user types Stop
                if (userQuestion.Trim().Equals("Stop", StringComparison.OrdinalIgnoreCase))
                    break;

                // Define the Excel assistant behavior
                string systemPrompt = @"
                You are an intelligent Excel assistant built using Syncfusion XlsIO and OpenAI. 
                Your job is to help users understand, analyze, and interact with Excel data through 
                natural language. 
                You are integrated into a C# application and can access structured Excel content 
                such as tables, charts, formulas, and cell values.
                Your capabilities include:
                - Summarizing Excel sheets and ranges
                - Answering questions about data trends, totals, averages, and comparisons
                - Explaining formulas and calculations
                - Suggesting improvements or insights based on the data
                - Responding in a friendly, professional, and concise manner
                Constraints:
                You do not generate or modify Excel files directly. 
                Instead, you interpret and explain the data provided to you by the application.
                Always assume the user is referring to the most recent Excel content unless stated 
                otherwise. 
                Maintain context across the conversation to provide coherent and helpful responses.
                If the user asks something outside the scope of Excel data, politely redirect 
                them back to Excel-related tasks.";

                // Build the user prompt
                string userPrompt = "Question:\n" + userQuestion;

                try
                {
                    // Generate a response for the user's question
                    string answer = await AskAzureOpenAIAsync(azureOpenAIApiKey, "your-model-name", systemPrompt, userPrompt, chatHistory);

                    // Display the chatbot response
                    Console.WriteLine("ChatBot: " + answer);
                }
                catch (Exception ex)
                {
                    // Display Azure OpenAI request errors
                    Console.WriteLine($"Azure OpenAI error: {ex.Message}");
                }
            }
        }

        // Sends a chat request to Azure OpenAI and returns the response
        private static async Task<string> AskAzureOpenAIAsync(string apiKey, string model, string systemPrompt, string userContent, List<ChatMessage> chatHistory)
        {
            // Create an Azure OpenAI client
            AzureOpenAIClient azureClient = new(
                new Uri("https://your-resource-name.openai.azure.com/"),
                new ApiKeyCredential(apiKey));

            // Create a chat client for the specified model
            ChatClient chatClient = azureClient.GetChatClient(model);

            // Add the system prompt to the conversation history
            chatHistory.Add(new SystemChatMessage(systemPrompt));

            // Add the user's input to the conversation history
            chatHistory.Add(new UserChatMessage(userContent));

            // Send the request to Azure OpenAI
            ClientResult<ChatCompletion> chatResult = await chatClient.CompleteChatAsync(chatHistory);

            // Extract the generated response
            string response = chatResult.Value.Content[0].Text ?? string.Empty;

            // Save the assistant response in the conversation history
            chatHistory.Add(new AssistantChatMessage(response));

            // Return the generated response
            return response;
        }

        // Converts workbook data into text for AI processing
        private static string BuildCsvContext(string excelFilePath)
        {
            // Create an Excel engine instance
            using ExcelEngine excelEngine = new ExcelEngine();

            // Access the Excel application
            IApplication application = excelEngine.Excel;

            // Set the Excel version
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the Excel file
            using FileStream fileStream = File.OpenRead(excelFilePath);

            // Load the workbook
            IWorkbook workbook = application.Workbooks.Open(fileStream);

            // Create a string builder to store workbook content
            StringBuilder stringBuilder = new StringBuilder();

            // Add workbook information
            stringBuilder.AppendLine($"File: {excelFilePath}");
            stringBuilder.AppendLine($"Sheets: {workbook.Worksheets.Count}");
            stringBuilder.AppendLine();

            // Iterate through all worksheets
            foreach (IWorksheet worksheet in workbook.Worksheets)
            {
                // Create a stream for CSV export
                MemoryStream csvStream = new MemoryStream();

                // Export the worksheet to CSV format
                worksheet.SaveAs(csvStream, ",");

                // Convert CSV content into text
                string excelData = Encoding.UTF8.GetString(csvStream.ToArray());

                // Add the worksheet name
                stringBuilder.AppendLine("Sheet Name : " + worksheet.Name);

                // Add worksheet data
                stringBuilder.AppendLine();
                stringBuilder.AppendLine(excelData);
                stringBuilder.AppendLine();
            }

            // Return the combined workbook content
            return stringBuilder.ToString();
        }
    }
}
