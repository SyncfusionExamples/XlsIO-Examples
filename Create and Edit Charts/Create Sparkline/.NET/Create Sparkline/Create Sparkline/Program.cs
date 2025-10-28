using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Create_Sparkline
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

                //Add SparklineGroups
                ISparklineGroup sparklineGroup = sheet.SparklineGroups.Add();

                //Add SparkLineType
                sparklineGroup.SparklineType = SparklineType.Line;
                sparklineGroup.MarkersColor = Color.BlueViolet;

                //Add sparklines
                ISparklines sparklines = sparklineGroup.Add();
                IRange dataRange = sheet.Range["D6:G17"];
                IRange referenceRange = sheet.Range["H6:H17"];
                sparklines.Add(dataRange, referenceRange);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Sparklines.xlsx"));
                #endregion
            }
        }
    }
}





