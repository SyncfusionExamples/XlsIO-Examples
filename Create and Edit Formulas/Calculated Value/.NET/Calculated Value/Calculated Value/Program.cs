using System.IO;
using Syncfusion.XlsIO;

namespace Calculated_Value
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

                sheet.Range["A1"].Value = "10";
                sheet.Range["B1"].Value = "20";

                sheet.Range["C1"].Formula = "=A1+B1";

                #region Calculated Value
                sheet.EnableSheetCalculations();

                //Returns the calculated value of a formula using the most current inputs
                string calculatedValue = sheet["C1"].CalculatedValue;
                sheet.Range["C3"].Value = "Calculated Value of the formula in C1 calculated through XlsIO is : " + calculatedValue;
                
                sheet.DisableSheetCalculations();
                #endregion

                sheet.Range["C3"].AutofitColumns();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Formula.xlsx"), FileMode.Create, FileAccess.Write);
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
