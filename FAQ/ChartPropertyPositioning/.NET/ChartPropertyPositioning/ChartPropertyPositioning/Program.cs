using Syncfusion.XlsIO;

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

            //Create a Chart
            IChartShape chart = sheet.Charts.Add();

            //Set Chart Type
            chart.ChartType = ExcelChartType.Column_Clustered;

            //Set data range in the worksheet
            chart.DataRange = sheet.Range["A1:C6"];
            chart.IsSeriesInRows = false;

            //Set Datalabels
            IChartSerie serie1 = chart.Series[0];
            IChartSerie serie2 = chart.Series[1];

            serie1.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
            serie2.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
            serie1.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;
            serie2.DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;

            //Set Legend
            chart.HasLegend = true;
            chart.Legend.Position = ExcelLegendPosition.Bottom;
            chart.ChartTitle = "Sales Analysis";

            //Positioning the chart in the worksheet
            chart.TopRow = 8;
            chart.LeftColumn = 1;
            chart.BottomRow = 23;
            chart.RightColumn = 8;

            //left positioning the legend
            chart.Legend.Position = ExcelLegendPosition.Left;

            //Manual layout positioning of the chart title
            chart.ChartTitleArea.Layout.LeftMode = LayoutModes.edge;
            chart.ChartTitleArea.Layout.TopMode = LayoutModes.edge;
            chart.ChartTitleArea.Layout.Left = 100;
            chart.ChartTitleArea.Layout.Top = 20;

            #region Save
            //Saving the workbook
            workbook.SaveAs(Path.GetFullPath("Output/Chart.xlsx"));
            #endregion
        }
    }
}
