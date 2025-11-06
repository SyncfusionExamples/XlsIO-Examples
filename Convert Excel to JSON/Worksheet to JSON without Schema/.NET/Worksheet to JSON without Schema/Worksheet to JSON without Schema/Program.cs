using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Worksheet_to_JSON_without_Schema
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

                #region save as JSON
                //Saves the workbook to JSON, as schema by default
                workbook.SaveAsJson(Path.GetFullPath("Output/Excel-Worksheet-To-JSON-filestream-without-schema.json"), worksheet,false);
                #endregion

                #region Open JSON 
                //Open default JSON
                #endregion
            }
        }
    }
}





