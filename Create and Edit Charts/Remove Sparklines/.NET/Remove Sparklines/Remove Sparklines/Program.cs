using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Sparklines
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

                ISparklineGroup sparklineGroup = sheet.SparklineGroups[0];
                ISparklines sparklines = sparklineGroup[0];

                //Remove sparkline specified by index from the sparklines
                sparklines.Remove(sparklines[1]);

                //Remove sparklines from the sparkline group
                sparklineGroup.Remove(sparklines);

                //Remove sparkline group from the sheet
                sheet.SparklineGroups.Remove(sparklineGroup);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveSparklines.xlsx"));
                #endregion
            }
        }
    }
}





