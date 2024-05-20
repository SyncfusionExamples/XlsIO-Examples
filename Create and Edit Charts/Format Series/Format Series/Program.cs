using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Format_Series
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Set name to chart series.
                chart.Series[0].Name = "Amount";

                //Set the series type.
                chart.Series[0].SerieType = ExcelChartType.Line_Markers;
                chart.Series[1].SerieType = ExcelChartType.Bar_Clustered;

                // Configure the fill settings for the series in the chart.
                chart.Series[1].SerieFormat.Fill.FillType = ExcelFillType.Gradient;
                chart.Series[1].SerieFormat.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chart.Series[1].SerieFormat.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chart.Series[1].SerieFormat.Fill.ForeColor = Color.Red;

                //Customize series border.
                chart.Series[1].SerieFormat.LineProperties.LineColor = Color.Red;
                chart.Series[1].SerieFormat.LineProperties.LinePattern = ExcelChartLinePattern.Dot;
                chart.Series[1].SerieFormat.LineProperties.LineWeight = ExcelChartLineWeight.Narrow;

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