using System.IO;
using Syncfusion.XlsIO;

namespace Access_Used_Range
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //UsedRange excludes the blank cell which has formatting
                worksheet.UsedRangeIncludesFormatting = false;

                #region Access UsedRange
                //Modifying the column width and row height of the used range
                worksheet.UsedRange.ColumnWidth = 20;
                worksheet.UsedRange.RowHeight = 20;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("AccessUsedRange.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("AccessUsedRange.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
