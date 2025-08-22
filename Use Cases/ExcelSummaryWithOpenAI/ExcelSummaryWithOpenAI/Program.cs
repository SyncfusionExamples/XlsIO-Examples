using OpenAI.Chat;
using Syncfusion.XlsIO;
using System.Text;


namespace ExcelSummaryWithOpenAI
{
    class Program
    {
        static async Task Main()
        {
            string excelFilePath = "../../../Data/Sales Data.xlsx"; // Replace with the path to your Excel file
            string openAiApiKey = "OpenAI API key"; //Replace with OpenAI API key
            try
            {
                // Read Excel data
                string excelText = ExtractDataAsText(excelFilePath);
                Console.WriteLine("Excel data read successfully.");

                // Get summary from OpenAI
                string summary = await SummarizeData(excelText, openAiApiKey);
                Console.WriteLine("Summary from OpenAI:");
                Console.WriteLine(summary);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        /// <summary>
        /// Extracts data from an Excel file and returns it as a formatted text string.
        /// </summary>
        /// <param name="filePath">Excel file path</param>
        /// <returns>Excel data as text</returns>
        public static string ExtractDataAsText(string filePath)
        {
            //Initialize the Excel engine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Get the Excel application instance
                IApplication application = excelEngine.Excel;

                // Set the default version to Xlsx
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Open the Excel file
                IWorkbook workbook = application.Workbooks.Open(filePath);

                // Create a StringBuilder to hold the extracted data
                StringBuilder sb = new StringBuilder();

                // Iterate through each worksheet and extract data
                foreach (IWorksheet worksheet in workbook.Worksheets)
                {
                    // Append the worksheet name
                    sb.AppendLine($"Worksheet: {worksheet.Name}");

                    // Get the used range of the worksheet
                    IRange usedRange = worksheet.UsedRange;


                    for (int row = usedRange.Row; row <= usedRange.LastRow; row++)
                    {
                        for (int col = usedRange.Column; col <= usedRange.LastColumn; col++)
                        {
                            //Get the cell value and append it to the StringBuilder
                            sb.Append(worksheet[row,col].DisplayText + "\t");
                        }
                        sb.AppendLine();
                    }

                    sb.AppendLine(); 
                }
                return sb.ToString();
            }
        }
        /// <summary>
        ///  Summarizes the provided text using OpenAI's GPT-5 model.
        /// </summary>
        /// <param name="inputText"> Input Excel data </param>
        /// <param name="apiKey">Open AI API key</param>
        /// <returns>AI Summary</returns>
        public static async Task<string> SummarizeData(string inputText, string apiKey)
        {
            // Initialize the OpenAI client with the provided API key and model
            ChatClient client = new(model: "gpt-5", apiKey);

            // Create a chat completion request to summarize the Excel data
            ChatCompletion completion = await client.CompleteChatAsync(
                new SystemChatMessage("You are a helpful assistant that summarizes Excel data."),
                new UserChatMessage($"Summarize the following Excel data:\n{inputText}")
            );

            // Check if the completion has content and return the summary
            return completion.Content[0].Text?.Trim() ?? string.Empty;
        }
    }
}