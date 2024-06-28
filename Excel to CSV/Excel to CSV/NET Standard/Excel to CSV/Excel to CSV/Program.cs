using Syncfusion.XlsIO;

namespace Excel_to_CSV
{
    class program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];
                sheet.Range["A1:M20"].Text = "document";

                //Saving the sheet and workbook as streams
                FileStream sheetStream = new FileStream("Sample.csv", FileMode.Create, FileAccess.ReadWrite);
                sheet.SaveAs(sheetStream, ",");

                FileStream stream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                sheetStream.Dispose();

                //Dispose streams
                sheetStream.Dispose();
                stream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Sample.csv")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}