using System.IO;
using Syncfusion.XlsIO;

namespace Save_TextFile
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                worksheet.Range["A1:M20"].Text = "Text document";

                #region Save as text file
                //Saving the workbook
                FileStream outputStream = new FileStream("TextFile.txt", FileMode.Create, FileAccess.Write);
                worksheet.SaveAs(outputStream, " ");
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("TextFile.txt")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
