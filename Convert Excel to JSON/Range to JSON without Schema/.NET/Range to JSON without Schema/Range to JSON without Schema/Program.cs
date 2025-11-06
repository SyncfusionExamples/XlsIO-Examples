using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Range_to_JSON_without_Schema
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
                workbook.SaveAsJson(Path.GetFullPath("Output/Excel-Range-To-JSON-filestream-without-schema.json"), range, false);
                #endregion

                #region Open JSON 
                //Open default JSON
                #endregion
            }
        }
    }
}





