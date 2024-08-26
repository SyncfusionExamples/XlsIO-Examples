using System.IO;
using Syncfusion.XlsIO;

namespace Hide_Row_and_Column_Headers
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
                IWorksheet sheet = workbook.Worksheets[0];

                sheet.Range["A1:M20"].Text = "RowColumnHeader";

                #region Hide Row and Column Headers
                sheet.IsRowColumnHeadersVisible = false;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/HideRowandColumnHeaders.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("HideRowandColumnHeaders.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
