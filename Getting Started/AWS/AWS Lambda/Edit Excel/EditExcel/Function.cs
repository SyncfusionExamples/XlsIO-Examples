using Amazon.Lambda.Core;
using Syncfusion.XlsIO;
using System.Reflection.Metadata;
using static System.Net.Mime.MediaTypeNames;

// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace EditExcel;

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
        //New instance of ExcelEngine is created 
        //Equivalent to launching Microsoft Excel with no workbooks open
        //Instantiate the spreadsheet creation engine
        ExcelEngine excelEngine = new ExcelEngine();

        //Instantiate the Excel application object
        IApplication application = excelEngine.Excel;

        //Assigns default application version
        application.DefaultVersion = ExcelVersion.Xlsx;

        //A existing workbook is opened.             
        FileStream sampleFile = new FileStream(@"Data/InputTemplate.xlsx", FileMode.Open);
        IWorkbook workbook = application.Workbooks.Open(sampleFile);

        //Access first worksheet from the workbook.
        IWorksheet worksheet = workbook.Worksheets[0];

        //Set Text in cell A3.
        worksheet.Range["A3"].Text = "Hello World";

        //Creating stream object.
        MemoryStream stream = new MemoryStream();
        workbook.SaveAs(stream);
        return Convert.ToBase64String(stream.ToArray());
    }
}
