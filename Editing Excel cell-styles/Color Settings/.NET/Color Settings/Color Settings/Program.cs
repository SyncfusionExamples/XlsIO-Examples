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
                workbook.SaveAs(Path.GetFullPath("Output/ColorSettings.xlsx"));
                #endregion
            }
        }
    }
}




