using System.IO;
using Syncfusion.XlsIO;

namespace Incremental_Formula
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Enables the incremental formula to updates the reference in cell
                application.EnableIncrementalFormula = true;

                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Formula are automatically increments by one for the range of cells
                sheet["A1:A5"].Formula = "=B1+C1";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("IncrementalFormula.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("IncrementalFormula.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
