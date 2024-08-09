using System.IO;
using Syncfusion.XlsIO;

namespace Set_Formula
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

                //Setting values to the cells
                sheet.Range["A1"].Number = 10;
                sheet.Range["B1"].Number = 10;

                #region Set Formula
                //Setting formula in the cell
                sheet.Range["C1"].Formula = "=SUM(A1,B1)";
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Formula.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Formula.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
