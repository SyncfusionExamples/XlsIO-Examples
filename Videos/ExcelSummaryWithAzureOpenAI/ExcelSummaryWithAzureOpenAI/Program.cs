using Azure;
using Azure.AI.OpenAI;
using OpenAI.Chat;
using Syncfusion.XlsIO;
using System.Text;

namespace ExcelSummaryWithAzureOpenAI
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            string excelFilePath = "../../../Data/Sales Data.xlsx"; // Replace with the path to your Excel file

            try
            {
                // Read Excel data and convert to text
                string excelText = ExtractDataAsText(excelFilePath);
                Console.WriteLine("Excel data read successfully.");

                // Send data to Azure OpenAI for summarization
                string summary = await SummarizeData(excelText);
                Console.WriteLine("Summary from OpenAI:");
                Console.WriteLine(summary);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // This method sends the extracted Excel text to Azure OpenAI
        public static Task<string> SummarizeData(string inputText)
        {
            // Initialize Azure OpenAI client with endpoint and API key
            AzureOpenAIClient azureOpenAIClient = new(
                new Uri("YOUR_ENDPOINT"),
                new AzureKeyCredential("YOUR_API_KEY"));

            // Create Chat client using the Azure OpenAI
            ChatClient chatClient = azureOpenAIClient.GetChatClient("YOUR_MODEL_NAME");

            // Create a chat completion request to summarize the Excel data
            ChatCompletion completion = chatClient.CompleteChat([
                new SystemChatMessage("You are a helpful assistant that summarizes Excel data."),
                new UserChatMessage($"Summarize the following Excel data:\n{inputText}")
                ]);

            // Return the summarized text result
            return Task.FromResult(completion.Content[0].Text?.Trim() ?? string.Empty);
        }

        // Extracts data from an Excel file and returns it as a formatted text string
        public static string ExtractDataAsText(string filePath)
        {
            //Initialize the Excel engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Get the Excel application instance
                IApplication application = excelEngine.Excel;

                // Set the default version to Xlsx
                application.DefaultVersion = ExcelVersion.Xlsx;

                // Open the Excel workbook
                IWorkbook workbook = application.Workbooks.Open(filePath);

                // Create a StringBuilder to hold the extracted data
                StringBuilder stringBuilder = new StringBuilder();

                // Iterate through each worksheet and extract data
                foreach (IWorksheet worksheet in workbook.Worksheets)
                {
                    // Append the worksheet name
                    stringBuilder.AppendLine($"Worksheet: {worksheet.Name}");

                    // Get the used range of the worksheet
                    IRange usedRange = worksheet.UsedRange;

                    // Loop through rows
                    for (int row = usedRange.Row; row <= usedRange.LastRow; row++)
                    {
                        // Loop through columns
                        for (int col = usedRange.Column; col <= usedRange.LastColumn; col++)
                        {
                            //Get the cell value and append it to the StringBuilder
                            stringBuilder.Append(worksheet[row, col].DisplayText + "\t");

                        }

                        stringBuilder.AppendLine();

                    }

                    stringBuilder.AppendLine();
                }

                return stringBuilder.ToString();
            }
        }
    }
}
