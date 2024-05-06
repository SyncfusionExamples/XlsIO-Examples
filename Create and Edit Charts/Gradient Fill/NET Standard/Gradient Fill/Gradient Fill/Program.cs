using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Charts;
using Syncfusion.XlsIO.Implementation.Shapes;
using Syncfusion.Drawing;
using Syncfusion.XlsIO.Implementation;

namespace Create_Chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a Chart
                IChartShape chart = sheet.Charts.Add();

                //Set Chart Type
                chart.ChartType = ExcelChartType.Column_Clustered;

                //Set data range in the worksheet
                chart.DataRange = sheet.Range["A1:C6"];
                chart.IsSeriesInRows = false;

                //Get Serie
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];

                //Set Datalabels
                serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie1.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;
                serie2.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;

                //Gradient fill for serie1
                ChartFillImpl chartFillImpl1 = serie1.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl1.FillType = ExcelFillType.Gradient;
                chartFillImpl1.GradientColorType = ExcelGradientColor.MultiColor;
                serie1.SerieFormat.Fill.GradientStyle = ExcelGradientStyle.Horizontal;
                GradientStopImpl gradientStopImpl1 = new GradientStopImpl(new ColorObject(Color.FromArgb(0, 176, 240)), 50000, 100000);
                GradientStopImpl gradientStopImpl2 = new GradientStopImpl(new ColorObject(Color.FromArgb(0, 112, 192)), 70000, 100000);
                chartFillImpl1.GradientStops.GradientType = GradientType.Liniar;
                chartFillImpl1.GradientStops.Add(gradientStopImpl1);
                chartFillImpl1.GradientStops.Add(gradientStopImpl2);

                //Gradient fill for serie2
                ChartFillImpl chartFillImpl2 = serie2.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl2.FillType = ExcelFillType.Gradient;
                chartFillImpl2.GradientColorType = ExcelGradientColor.MultiColor;
                serie2.SerieFormat.Fill.GradientStyle = ExcelGradientStyle.Horizontal;
                GradientStopImpl gradientStopImpl3 = new GradientStopImpl(new ColorObject(Color.FromArgb(244, 177, 131)), 40000, 100000);
                GradientStopImpl gradientStopImpl4 = new GradientStopImpl(new ColorObject(Color.FromArgb(255, 102, 0)), 70000, 100000);
                chartFillImpl2.GradientStops.GradientType = GradientType.Liniar;
                chartFillImpl2.GradientStops.Add(gradientStopImpl3);
                chartFillImpl2.GradientStops.Add(gradientStopImpl4);

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                //Positioning the chart in the worksheet
                chart.TopRow = 8;
                chart.LeftColumn = 1;
                chart.BottomRow = 23;
                chart.RightColumn = 8;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("GradientFill.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("GradientFill.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
