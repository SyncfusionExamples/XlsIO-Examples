using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Format_Plot_Area
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Format Plot Area
                IChartFrameFormat chartPlotArea = chart.PlotArea;

                //Set border line pattern, color, line weight.
                chartPlotArea.Border.LinePattern = ExcelChartLinePattern.Solid;
                chartPlotArea.Border.LineColor = Color.Blue;
                chartPlotArea.Border.LineWeight = ExcelChartLineWeight.Hairline;

                //Set fill type and color.
                chartPlotArea.Fill.FillType = ExcelFillType.Gradient;
                chartPlotArea.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chartPlotArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartPlotArea.Fill.ForeColor = Color.White;

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}