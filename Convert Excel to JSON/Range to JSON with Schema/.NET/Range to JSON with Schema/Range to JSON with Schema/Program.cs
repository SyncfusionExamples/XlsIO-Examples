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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Custom range
                IRange range = worksheet.Range["A1:F100"];

                #region save as JSON
                //Saves the workbook to a JSON filestream, as schema by default
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Excel-Range-To-JSON-as-schema-default.json"), FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAsJson(outputStream, range);

                //Saves the workbook to a JSON filestream as schema
                FileStream stream1 = new FileStream("Excel-Range-To-JSON-as-schema.json", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAsJson(stream1, range, true);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                stream1.Dispose();
                inputStream.Dispose();

                #region Open JSON 
                //Open default JSON

                //Open JSON with Schema
                #endregion
            }
        }
    }
}





