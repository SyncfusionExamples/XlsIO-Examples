using Amazon.Lambda.Core;
using static System.Net.Mime.MediaTypeNames;
using Syncfusion.XlsIO;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace LoadingandSaving;

public class Function
{
    
    /// <summary>
    /// A simple function that takes a string and does a ToUpper
    /// </summary>
    /// <param name="input">The event for the Lambda function handler to process.</param>
    /// <param name="context">The ILambdaContext that provides methods for logging and describing the Lambda environment.</param>
    /// <returns></returns>
    public string FunctionHandler(string input, ILambdaContext context)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Data", "InputTemplate.xlsx");
            FileStream excelStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(excelStream);

            //Access first worksheet
            IWorksheet worksheet = workbook.Worksheets[0];

            //Set text in cell A3
            worksheet.Range["A3"].Text = "Hello World";

            //Save to MemoryStream
            MemoryStream outputStream = new MemoryStream();
            workbook.SaveAs(outputStream);
            outputStream.Position = 0;

            //Return as Base64 string
            return Convert.ToBase64String(outputStream.ToArray());
        }
    }
}
