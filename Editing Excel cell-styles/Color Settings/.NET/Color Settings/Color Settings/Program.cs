using System.IO;
using Syncfusion.XlsIO;

namespace Color_Settings
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

                #region Color Settings
                //Apply cell back color
                worksheet.Range["A1"].CellStyle.ColorIndex = ExcelKnownColors.Aqua;

                //Apply cell pattern
                worksheet.Range["A2"].CellStyle.FillPattern = ExcelPattern.Angle;

                //Apply cell fore color
                worksheet.Range["A2"].CellStyle.PatternColorIndex = ExcelKnownColors.Green;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ColorSettings.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ColorSettings.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
