using System.IO;
using Syncfusion.XlsIO;

namespace Hide_Gridlines
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
                sheet.Range["A1:M20"].Text = "Gridlines";

                #region Hide Gridlines
                //Hide grid line
                sheet.IsGridLinesVisible = false;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/HideGridlines.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("HideGridlines.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
