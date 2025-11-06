using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Range_to_JSON_with_Schema
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

                //Custom range
                IRange range = worksheet.Range["A1:F100"];

                #region save as JSON
                //Saves the workbook to JSON, as schema by default
                workbook.SaveAsJson(Path.GetFullPath("Output/Excel-Range-To-JSON-as-schema-default.json"), range);

                //Saves the workbook to JSON as schema
                workbook.SaveAsJson(Path.GetFullPath("Output/Excel-Range-To-JSON-as-schema.json"), range, true);
                #endregion

                #region Open JSON 
                //Open default JSON

                //Open JSON with Schema
                #endregion
            }
        }
    }
}





