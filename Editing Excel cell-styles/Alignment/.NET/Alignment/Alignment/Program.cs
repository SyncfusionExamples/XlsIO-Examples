using System.IO;
using Syncfusion.XlsIO;

namespace Alignment
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

                worksheet.Range["A2"].Text = "HAlignCenter";
                worksheet.Range["A4"].Text = "HAlignFill";
                worksheet.Range["A6"].Text = "HAlignRight";
                worksheet.Range["A8"].Text = "HAlignCenterAcrossSelection";
                worksheet.Range["B2"].Text = "VAlignCenter";
                worksheet.Range["B4"].Text = "VAlignFill";
                worksheet.Range["B6"].Text = "VAlignTop";
                worksheet.Range["B8"].Text = "VAlignCenterAcrossSelection";
                worksheet.Range["C2"].Text = "Text Rotation to 60 degree";
                worksheet.Range["C4"].Text = "Text Rotation to 90 degree";
                worksheet.Range["C6"].Text = "Indent level is 6";
                worksheet.Range["D2"].Text = "Text Direction(LeftToRight)";
                worksheet.Range["D3"].Text = "Text Direction(RightToLeft)";
                worksheet.Range["D4"].Text = "Text Direction(Context)";

                #region Alignment
                //Text Alignment Setting (Horizontal Alignment)
                worksheet.Range["A2"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                worksheet.Range["A4"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignFill;
                worksheet.Range["A6"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;
                worksheet.Range["A8"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenterAcrossSelection;

                //Text Alignment Setting (Vertical Alignment)
                worksheet.Range["B2"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignBottom;
                worksheet.Range["B4"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignCenter;
                worksheet.Range["B6"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignTop;
                worksheet.Range["B8"].CellStyle.VerticalAlignment = ExcelVAlign.VAlignDistributed;

                //Text Orientation Settings
                worksheet.Range["C2"].CellStyle.Rotation = 60;
                worksheet.Range["C4"].CellStyle.Rotation = 90;

                //Text Indent Setting
                worksheet.Range["C6"].CellStyle.IndentLevel = 6;

                //Text Direction Setting
                worksheet.Range["D2"].CellStyle.ReadingOrder = ExcelReadingOrderType.LeftToRight;
                worksheet.Range["D3"].CellStyle.ReadingOrder = ExcelReadingOrderType.RightToLeft;
                worksheet.Range["D4"].CellStyle.ReadingOrder = ExcelReadingOrderType.Context;
                #endregion

                worksheet.UsedRange.AutofitColumns();
                worksheet.UsedRange.AutofitRows();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Alignment.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Alignment.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
