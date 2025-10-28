using System.IO;
using Syncfusion.XlsIO;

namespace Edit_Sparkline
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Edit Sparklines
                ISparklineGroup sparklineGroup = sheet.SparklineGroups[0];
                ISparklines sparklines = sparklineGroup[0];
                IRange dataRange = sheet["D6:F17"];
                IRange referenceRange = sheet["H6:H17"];

                //Edit the existing sparklines data
                sparklines.RefreshRanges(dataRange, referenceRange);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/EditSparklines.xlsx"));
                #endregion
            }
        }
    }
}





