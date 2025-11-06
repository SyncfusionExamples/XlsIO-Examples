using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Workbook_to_JSON_with_Schema
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Saves the workbook to JSON as schema
                workbook.SaveAsJson(Path.GetFullPath("Output/Excel-Workbook-To-JSON-as-schema.json"), true);
            }
        }
    }
}