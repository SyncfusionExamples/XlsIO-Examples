using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Chart_Area
{
    class program
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

                //Format Chart Area
                IChartFrameFormat chartArea = chart.ChartArea;

                //Set the border
                chartArea.Border.LinePattern = ExcelChartLinePattern.Solid;
                chartArea.Border.LineColor = Color.Blue;
                chartArea.Border.LineWeight = ExcelChartLineWeight.Hairline;

                //Set the color
                chartArea.Fill.FillType = ExcelFillType.Gradient;
                chartArea.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chartArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Color.White;

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}




