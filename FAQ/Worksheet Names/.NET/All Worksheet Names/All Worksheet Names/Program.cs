using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;

namespace All_Worksheet_Names
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));

                //Get the worksheets collection
                WorksheetsCollection worksheets = workbook.Worksheets as WorksheetsCollection;

                //Print all worksheet names
                foreach (IWorksheet worksheet in worksheets)
                {
                    Console.WriteLine(worksheet.Name);
                }
            }
        }
    }
}