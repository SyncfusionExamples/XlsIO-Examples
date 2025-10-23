using OpenAI;
using OpenAI.Chat;
using Syncfusion.XlsIO;
using System;
using System.ClientModel;
using System.Linq;
using System.Text;

/// <summary>
///  AI-powered Excel translator for using Syncfusion XlsIO and OpenAI.
/// </summary>
class ExcelContentTranslator
{
    /// <summary>
    /// Entry point for the Translator.
    /// </summary>
    static async Task Main()
    {
        // Replace with your actual OpenAI API key or set it in environment variables for security
        string? openAIApiKey = "Replace the OpenAI Key";
        await ExecuteTranslation(openAIApiKey);
    }

    /// <summary>
    /// Execute translation of Excel document.
    /// </summary>
    private async static Task ExecuteTranslation(string openAIApiKey)
    {
        Console.WriteLine("AI Powered Excel Translator");

        Console.WriteLine("Enter full Excel file path (e.g., C:\\Data\\report.xlsx):");

        //Read user input for Excel file path
        string? excelFilePath = Console.ReadLine()?.Trim().Trim('"');

        Console.WriteLine("Enter the language name (e.g., Chinese, Japanese)");

        //Read user input for required language
        string? language = Console.ReadLine()?.Trim().Trim('"');

        if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
        {
            Console.WriteLine("Invalid path. Exiting.");
            return;
        }

        if (string.IsNullOrWhiteSpace(language))
        {
            Console.WriteLine("Invalid language. Exiting.");
            return;
        }


        if (string.IsNullOrWhiteSpace(openAIApiKey))
        {
            Console.WriteLine("OPENAI_API_KEY not set. Exiting.");
            return;
        }

        try
        {
            // Translate Excel content
            await TranslateExcelContent(openAIApiKey, excelFilePath, language);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to read Excel: {ex.StackTrace}");
            return;
        }
    }

    /// <summary>
    /// Translate Excel content using OpenAI and Syncfusion XlsIO.
    /// </summary>
    /// <param name="openAIApiKey">OpenAI API key.</param>
    /// <param name="excelFilePath">Path to the Excel file.</param>
    /// <param name="language">Target language for translation.</param>
    private static async Task TranslateExcelContent(string openAIApiKey, string excelFilePath, string language)
    {
        // Initialize Syncfusion Excel engine
        using ExcelEngine excelEngine = new ExcelEngine();

        // Set default version to XLSX
        IApplication excelApp = excelEngine.Excel;
        excelApp.DefaultVersion = ExcelVersion.Xlsx;

        IWorkbook workbook = excelApp.Workbooks.Open(excelFilePath);

        // Store translation results
        Dictionary<string, string> translationResults = new Dictionary<string, string>();

        // Convert each worksheet to CSV format and append to the context
        foreach (IWorksheet worksheet in workbook.Worksheets)
        {
            worksheet.UsedRangeIncludesFormatting = false;
            IRange usedRange = worksheet.UsedRange;

            int firstRow = usedRange.Row;
            int lastRow = usedRange.LastRow;
            int firstCol = usedRange.Column;
            int lastCol = usedRange.LastColumn;

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


            for (int row = firstRow; row <= lastRow; row++)
            {
                for (int col = firstCol; col <= lastCol; col++)
                {

                    // Skip formula, number, boolean, and date time cells
                    if (worksheet[row, col].HasBoolean || worksheet[row, col].HasDateTime || worksheet[row, col].HasFormula || worksheet[row, col].HasNumber)
                    {
                        continue;
                    }

                    // Get cell value
                    string cellValue = worksheet.GetCellValue(row, col, false);

                    // Skip empty cells
                    if (string.IsNullOrEmpty(cellValue))
                    {
                        continue;
                    }

                    // Prepare user prompt
                    string userPrompt = cellValue;                    

                    try
                    {
                        string translatedText = cellValue;
                        
                        if (!translationResults.TryGetValue(cellValue, out translatedText))
                        {
                            // Get translated text from OpenAI
                            translatedText = await AskOpenAIAsync(openAIApiKey, "gpt-4o-mini", systemPrompt, userPrompt);
                            translationResults.Add(cellValue, translatedText);
                        }
                        worksheet.SetValue(row, col, translatedText);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"OpenAI error: {ex.Message}");
                    }
                }

            }
        }
        workbook.SaveAs("TranslatedExcelDocument.xlsx");
    }

    /// <summary>
    /// Sends a chat completion request to OpenAI and returns the response.
    /// </summary>
    /// <param name="apiKey">OpenAI API key.</param>
    /// <param name="model">Model name.</param>
    /// <param name="systemPrompt">System prompt.</param>
    /// <param name="userContent">User content.</param>
    /// <returns>AI-generated response as a string.</returns>
    private static async Task<string> AskOpenAIAsync(string apiKey, string model, string systemPrompt, string userContent)
    {
        // Initialize OpenAI client
        OpenAIClient openAIClient = new OpenAIClient(apiKey);

        // Create chat client for the specified model
        ChatClient chatClient = openAIClient.GetChatClient(model);

        //Get AI response
        ClientResult<ChatCompletion> chatResult = await chatClient.CompleteChatAsync(new SystemChatMessage(systemPrompt), new UserChatMessage(userContent));

        string response = chatResult.Value.Content[0].Text ?? string.Empty;

        return response;
    }

}