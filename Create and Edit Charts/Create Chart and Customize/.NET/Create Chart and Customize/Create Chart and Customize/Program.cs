using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Create_Chart_and_Customize
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing Excel file
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create a Chart
                IChartShape chart = worksheet.Charts.Add();

                //Set the Chart Type
                chart.ChartType = ExcelChartType.Column_Clustered;

                //Set data range in the worksheet
                chart.DataRange = worksheet.Range["A1:C6"];

                //Specify that the series are in columns
                chart.IsSeriesInRows = false;

                //Positioning the chart in the worksheet
                chart.TopRow = 8;
                chart.LeftColumn = 1;
                chart.BottomRow = 23;
                chart.RightColumn = 8;

                //Set the chart title
                chart.ChartTitle = "Purchase Details";

                //Format chart title color and font
                chart.ChartTitleArea.Color = ExcelKnownColors.Black;

                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Bold = true;
                chart.ChartTitleArea.Underline = ExcelUnderline.Single;
                chart.ChartTitleArea.Size = 15;

                //Format Chart Area
                IChartFrameFormat chartArea = chart.ChartArea;

                //Format chart area border and color
                chartArea.Border.LinePattern = ExcelChartLinePattern.Solid;
                chartArea.Border.LineColor = Color.Pink;
                chartArea.Border.LineWeight = ExcelChartLineWeight.Hairline;

                chartArea.Fill.FillType = ExcelFillType.Gradient;
                chartArea.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chartArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Color.White;

                //Format Plot Area
                IChartFrameFormat chartPlotArea = chart.PlotArea;

                //Format plot area border and color
                chartPlotArea.Border.LinePattern = ExcelChartLinePattern.Solid;
                chartPlotArea.Border.LineColor = Color.Pink;
                chartPlotArea.Border.LineWeight = ExcelChartLineWeight.Hairline;

                chartPlotArea.Fill.FillType = ExcelFillType.Gradient;
                chartPlotArea.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chartPlotArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartPlotArea.Fill.ForeColor = Color.White;

                //Format Series
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];

                //Format series border and color
                serie1.SerieFormat.LineProperties.LineColor = Color.Pink;
                serie1.SerieFormat.LineProperties.LinePattern = ExcelChartLinePattern.Dot;
                serie1.SerieFormat.LineProperties.LineWeight = ExcelChartLineWeight.Narrow;

                serie2.SerieFormat.LineProperties.LineColor = Color.Pink;
                serie2.SerieFormat.LineProperties.LinePattern = ExcelChartLinePattern.Dot;
                serie2.SerieFormat.LineProperties.LineWeight = ExcelChartLineWeight.Narrow;

                serie1.SerieFormat.Fill.FillType = ExcelFillType.Gradient;
                serie1.SerieFormat.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                serie1.SerieFormat.Fill.BackColor = Color.FromArgb(205, 217, 234);
                serie1.SerieFormat.Fill.ForeColor = Color.Pink;

                serie2.SerieFormat.Fill.FillType = ExcelFillType.Gradient;
                serie2.SerieFormat.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                serie2.SerieFormat.Fill.BackColor = Color.FromArgb(205, 217, 234);
                serie2.SerieFormat.Fill.ForeColor = Color.Pink;

                //Set Datalabel
                serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                serie1.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;
                serie2.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;

                //Format data labels color and font
                serie1.DataPoints.DefaultDataPoint.DataLabels.Color = ExcelKnownColors.Black;
                serie2.DataPoints.DefaultDataPoint.DataLabels.Color = ExcelKnownColors.Black;

                serie1.DataPoints.DefaultDataPoint.DataLabels.Size = 10;
                serie1.DataPoints.DefaultDataPoint.DataLabels.FontName = "calibri";
                serie1.DataPoints.DefaultDataPoint.DataLabels.Bold = true;

                serie2.DataPoints.DefaultDataPoint.DataLabels.Size = 10;
                serie2.DataPoints.DefaultDataPoint.DataLabels.FontName = "calibri";
                serie2.DataPoints.DefaultDataPoint.DataLabels.Bold = true;

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                //Format legend border, color, and font
                chart.Legend.FrameFormat.Border.AutoFormat = false;
                chart.Legend.FrameFormat.Border.IsAutoLineColor = false;
                chart.Legend.FrameFormat.Border.LineColor = Color.Black;
                chart.Legend.FrameFormat.Border.LinePattern = ExcelChartLinePattern.LightGray;
                chart.Legend.FrameFormat.Border.LineWeight = ExcelChartLineWeight.Narrow;

                chart.Legend.TextArea.Color = ExcelKnownColors.Black;

                chart.Legend.TextArea.Bold = true;
                chart.Legend.TextArea.FontName = "Calibri";
                chart.Legend.TextArea.Size = 8;
                chart.Legend.TextArea.Strikethrough = false;

                //Set axis title
                chart.PrimaryCategoryAxis.Title = "Items";
                chart.PrimaryValueAxis.Title = "Amount in($) and counts";

                //Format chart axis border and font
                chart.PrimaryCategoryAxis.Border.LinePattern = ExcelChartLinePattern.CircleDot;
                chart.PrimaryCategoryAxis.Border.LineColor = Color.Pink;
                chart.PrimaryCategoryAxis.Border.LineWeight = ExcelChartLineWeight.Hairline;

                chart.PrimaryValueAxis.Border.LinePattern = ExcelChartLinePattern.CircleDot;
                chart.PrimaryValueAxis.Border.LineColor = Color.Pink;
                chart.PrimaryValueAxis.Border.LineWeight = ExcelChartLineWeight.Hairline;

                chart.PrimaryCategoryAxis.Font.Color = ExcelKnownColors.Black;
                chart.PrimaryCategoryAxis.Font.FontName = "Calibri";
                chart.PrimaryCategoryAxis.Font.Bold = true;
                chart.PrimaryCategoryAxis.Font.Size = 8;

                chart.PrimaryValueAxis.Font.Color = ExcelKnownColors.Black;
                chart.PrimaryValueAxis.Font.FontName = "Calibri";
                chart.PrimaryValueAxis.Font.Bold = true;
                chart.PrimaryValueAxis.Font.Size = 8;

                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Chart.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose stream
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}
