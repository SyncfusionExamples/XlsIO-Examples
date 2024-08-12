using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace WaterMark
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing Excel file
                FileStream inputStream = new FileStream(@"../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Insert image in the worksheet for watermark
                FileStream imageStream = new FileStream(@"../../../Data/Watermark.png", FileMode.Open, FileAccess.Read);
                worksheet.PageSetup.BackgoundImage = new Image(imageStream);

                //Save the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                inputStream.Dispose();
                imageStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}