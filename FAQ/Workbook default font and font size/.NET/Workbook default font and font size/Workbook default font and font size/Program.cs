using System;
using System.IO;
using Syncfusion.XlsIO;

namespace WorkbookDefaultFontAndFontSize
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Set default font and font size
                workbook.StandardFont = "Calibri";
                workbook.StandardFontSize = 12;

                //Add some text
                sheet.Range["A1"].Text = "This is default font and size";

                //Save to file
                FileStream outputStream = new FileStream("Output/Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();

            }
        }
    }
}