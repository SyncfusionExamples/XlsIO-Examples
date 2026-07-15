using Syncfusion.XlsIO;
using Azure.AI.OpenAI;
using OpenAI.Chat;
using System.ClientModel;

namespace ExcelContentTranslator
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string? azureOpenAIApiKey = "Replace your Azure OpenAI key"; 
            await ExecuteTranslation(azureOpenAIApiKey);
        }

        // Collects user input, validates it, and initiates Excel translation
        private async static Task ExecuteTranslation(string azureOpenAIApiKey)
        {
            Console.WriteLine("AI Powered Excel Translator");
            Console.WriteLine("Enter full Excel file path (e.g., C:\\Data\\report.xlsx):");
            string? excelFilePath = Console.ReadLine()?.Trim().Trim('"');
            Console.WriteLine("Enter the language name (e.g., Chinese, Japanese)");
            string? language = Console.ReadLine()?.Trim().Trim('"');

            // Validate the Excel file path
            if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
            {
                Console.WriteLine("Invalid path. Exiting.");
                return;
            }

            // Validate the target language
            if (string.IsNullOrWhiteSpace(language))
            {
                Console.WriteLine("Invalid language. Exiting.");
                return;
            }

            // Validate the Azure OpenAI API key
            if (string.IsNullOrWhiteSpace(azureOpenAIApiKey))
            {
                Console.WriteLine("AZURE_OPENAI_API_KEY not set. Exiting.");
                return;
            }

            try
            {
                // Translate the workbook content
                await TranslateExcelContent(azureOpenAIApiKey, excelFilePath, language);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to read Excel: {ex.StackTrace}");
                return;
            }
        }

        // Translates the text content in the Excel workbook using Azure OpenAI
        private static async Task TranslateExcelContent(string azureOpenAIApiKey, string excelFilePath, string language)
        {
            // Initialize Syncfusion Excel engine
            using ExcelEngine excelEngine = new ExcelEngine();

            // Set default version to XLSX
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;


            IWorkbook workbook = application.Workbooks.Open(excelFilePath);
            Dictionary<string, string> translationResults = [];

            foreach (IWorksheet worksheet in workbook.Worksheets)
            {
                worksheet.UsedRangeIncludesFormatting = false;
                IRange usedRange = worksheet.UsedRange;

                int firstRow = usedRange.Row;
                int lastRow = usedRange.LastRow;
                int firstCol = usedRange.Column;
                int lastCol = usedRange.LastColumn;

                // Define translation instructions for Azure OpenAI
                string systemPrompt = @"You are a professional translator integrated into an Excel automation tool.
                                    Your job is to translate text from Excel cells into the" + language + @" language
                                    Rules:
                                    - Preserve original structure as much as possible.
                                    - Prefer literal translation over paraphrasing.
                                    - Return ONLY the translated text, without quotes, labels, or explanations.
                                    - Preserve placeholders (e.g., {0}, {name}) and keep numbers, currency, and dates intact.
                                    - Do not change the meaning, tone, or formatting unnecessarily.
                                    - Do not add extra commentary or code fences.
                                    - If the text is already in the target language, return it unchanged.
                                    - Be concise and accurate.";

                // Iterate through all cells
                for (int row = firstRow; row <= lastRow; row++)
                {
                    for (int col = firstCol; col <= lastCol; col++)
                    {
                        // Skip cells containing numbers, formulas, dates, or boolean values
                        if (worksheet[row, col].HasBoolean || worksheet[row, col].HasDateTime || worksheet[row, col].HasFormula || worksheet[row, col].HasNumber)
                        {
                            continue;
                        }

                        // Retrieve the cell text
                        string cellValue = worksheet.GetCellValue(row, col, false);

                        // Skip empty cells
                        if (string.IsNullOrEmpty(cellValue))
                        {
                            continue;
                        }

                        string userPrompt = cellValue;

                        try
                        {
                            string? translatedText = cellValue;

                            // Reuse translations for repeated text values
                            if (!translationResults.TryGetValue(cellValue, out translatedText))
                            {
                                // Send the cell content to Azure OpenAI for translation
                                translatedText = await AskAzureOpenAIAsync(azureOpenAIApiKey, systemPrompt, userPrompt);
                                translationResults.Add(cellValue, translatedText);
                            }

                            // Update the cell with the translated text
                            worksheet.SetValue(row, col, translatedText);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"OpenAI error: {ex.Message}");
                        }
                    }
                }
            }

            // Save the translated workbook to the desktop
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(path, "TranslatedExcelDocument.xlsx");
            workbook.SaveAs(filePath);
        }

        // This method sends the document text to Azure OpenAI and returns the translated result
        private static async Task<string> AskAzureOpenAIAsync(string apiKey, string systemPrompt, string userContent)
        {
            // Initialize the Azure OpenAI client with the endpoint and API key
            AzureOpenAIClient azureOpenAIClient = new(
                new Uri("https://your-resource-name.openai.azure.com/"),
                new ApiKeyCredential(apiKey)
                );

            // Create Chat client using the Azure OpenAI
            ChatClient chatClient = azureOpenAIClient.GetChatClient("your-model-name");

            // Send the translation request to Azure OpenAI
            ClientResult<ChatCompletion> chatResult = await chatClient.CompleteChatAsync(
                new SystemChatMessage(systemPrompt),
                new UserChatMessage(userContent)
                );
            string response = chatResult.Value.Content[0].Text ?? string.Empty;

            // Return the translated text
            return response; 
        }
    }
}
