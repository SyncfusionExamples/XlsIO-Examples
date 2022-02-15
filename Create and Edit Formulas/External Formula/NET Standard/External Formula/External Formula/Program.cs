using System.IO;
using Syncfusion.XlsIO;

namespace External_Formula
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

                //Write an external formula value
                sheet.Range["C1"].Formula = "[C:/Syncfusion/One.xlsx]Sheet1!$A$1*5";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ExternalFormula.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ExternalFormula.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
