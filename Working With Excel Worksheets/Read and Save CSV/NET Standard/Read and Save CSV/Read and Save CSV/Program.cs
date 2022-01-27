using System.IO;
using Syncfusion.XlsIO;

namespace Read_and_Save_CSV
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.csv", FileMode.Open, FileAccess.Read);

                #region Read CSV
                //Open the Tab delimited CSV file
                IWorkbook workbook = application.Workbooks.Open(inputStream, "\t");
                #endregion

                IWorksheet worksheet = workbook.Worksheets[0];

                #region Save CSV
                //Saving the workbook
                FileStream outputStream = new FileStream("ReadandSaveCSV.csv", FileMode.Create, FileAccess.Write);
                worksheet.SaveAs(outputStream, ",");
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ReadandSaveCSV.csv")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
