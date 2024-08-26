using System.IO;
using Syncfusion.XlsIO;

namespace Access_Discontinuous_Range
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

                #region Discontinuous Range
                //range1 and range2 are discontinuous ranges
                IRange range1 = sheet.Range["A1:A2"];
                IRange range2 = sheet.Range["C1:C2"];
                IRanges ranges = sheet.CreateRangesCollection();

                //range1 and range2 are considered as a single range
                ranges.Add(range1);
                ranges.Add(range2);
                ranges.Text = "Test";
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/DiscontinuousRange.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




