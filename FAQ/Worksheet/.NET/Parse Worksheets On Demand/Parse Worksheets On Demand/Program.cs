using System;
using System.IO;
using Syncfusion.XlsIO;



namespace Parse_Worksheets_On_Demand
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"), ExcelOpenType.Automatic, ExcelParseOptions.ParseWorksheetsOnDemand);

                // Access the first worksheet (triggers parsing)
                IWorksheet worksheet = workbook.Worksheets[0];

                // Process your data
                string value = worksheet.Range["A1"].Text;

                // Save to file system
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                workbook.Close();
                excelEngine.Dispose();
            }
        }
    }
}