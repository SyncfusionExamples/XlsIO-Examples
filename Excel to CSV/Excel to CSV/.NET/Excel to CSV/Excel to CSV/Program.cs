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
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Saving the workbook as streams
                FileStream outputStream = new FileStream("Sample.csv", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream, ",");

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

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