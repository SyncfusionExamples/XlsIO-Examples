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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
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
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/RemoveSparklines.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





