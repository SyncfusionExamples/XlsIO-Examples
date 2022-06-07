using System;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using System.IO;

namespace Table_To_Range
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/Sample.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Initialize XlsIO renderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document to make the table style apply to cells
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Get the worksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the table location
                IRange location = worksheet.ListObjects[0].Location;

                //Create a temp worksheet
                IWorksheet tempSheet = workbook.Worksheets.Create("TempSheet");

                //Copy the contents in table range in main worksheet to temp worksheet
                worksheet.Range[location.Row, location.Column, location.LastRow, location.LastColumn].CopyTo(tempSheet.Range[location.Row, location.Column], ExcelCopyRangeOptions.CopyValueAndSourceFormatting);
                
                //Clear the contents in table range in main worksheet
                worksheet.ListObjects[0].Location.Clear(ExcelClearOptions.ClearAll);

                //Copy the contents from temp worksheet to main worksheet
                tempSheet.Range[location.Row, location.Column, location.LastRow, location.LastColumn].CopyTo(worksheet.Range[location.Row, location.Column], ExcelCopyRangeOptions.CopyValueAndSourceFormatting);
                
                //Remove the temp worksheet
                tempSheet.Remove();

                //Remove borders
                worksheet.UsedRange.BorderNone();

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);
            }
        }
    }
}
