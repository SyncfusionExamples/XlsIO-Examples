using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Plot_Area
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Format Plot Area
                IChartFrameFormat chartPlotArea = chart.PlotArea;

                //Set the border
                chartPlotArea.Border.LinePattern = ExcelChartLinePattern.Solid;
                chartPlotArea.Border.LineColor = Color.Blue;
                chartPlotArea.Border.LineWeight = ExcelChartLineWeight.Hairline;

                //Set the color.
                chartPlotArea.Fill.FillType = ExcelFillType.Gradient;
                chartPlotArea.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chartPlotArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartPlotArea.Fill.ForeColor = Color.White;

                //Set the position
                chartPlotArea.Layout.Left = 5;

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}




