using Amazon.Lambda.Core;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using static System.Net.Mime.MediaTypeNames;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace Convert_Excel_to_Image;

public class Function
{
    
    /// <summary>
    /// A simple function that takes a string and does a ToUpper
    /// </summary>
    /// <param name="input"></param>
    /// <param name="context"></param>
    /// <returns></returns>
    public string FunctionHandler(string input, ILambdaContext context)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            //Initialize XlsIO renderer.
            application.XlsIORenderer = new XlsIORenderer();

            FileStream excelStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(excelStream);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Create the MemoryStream to save the image.      
            MemoryStream imageStream = new MemoryStream();

            //Save the converted image to MemoryStream.
            worksheet.ConvertToImage(worksheet.UsedRange, imageStream);
            imageStream.Position = 0;
            return Convert.ToBase64String(imageStream.ToArray());
        }
    }
}
