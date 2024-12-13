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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Saves the workbook to a JSON filestream as schema
                FileStream jsonWithSchema = new FileStream(Path.GetFullPath("Output/Excel-Workbook-To-JSON-as-schema.json"), FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAsJson(jsonWithSchema, true);

                //Dispose streams
                jsonWithSchema.Dispose();
                inputStream.Dispose();
            }
        }
    }
}
