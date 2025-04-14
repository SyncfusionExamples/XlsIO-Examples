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
                FileStream inputStream = new FileStream("Data/Input.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic, ExcelParseOptions.ParseWorksheetsOnDemand);

                // Access the first worksheet (triggers parsing)
                IWorksheet worksheet = workbook.Worksheets[0];

                // Process your data
                string value = worksheet.Range["A1"].Text;

                // Save to file system
                FileStream stream = new FileStream("Output/Output.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                workbook.Close();
                excelEngine.Dispose();
            }

        }

    }
}