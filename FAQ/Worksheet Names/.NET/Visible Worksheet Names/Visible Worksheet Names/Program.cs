using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;

namespace Visible_Worksheet_Names
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("Data/Input.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Get the worksheets collection
                WorksheetsCollection worksheets = workbook.Worksheets as WorksheetsCollection;

                //Print visible worksheet names
                foreach (IWorksheet worksheet in worksheets)
                {
                    if (worksheet.Visibility == WorksheetVisibility.Visible)
                        Console.WriteLine(worksheet.Name);
                }

                //Dispose streams
                inputStream.Dispose();

            }
        }
    }
}