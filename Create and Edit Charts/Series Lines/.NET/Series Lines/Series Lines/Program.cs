using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Series_Lines
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine engine = new ExcelEngine())
            {
                IApplication application = engine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a Chart
                IChartShape chart = sheet.Charts.Add();

                //Set Chart Type
                chart.ChartType = ExcelChartType.Bar_Stacked;

                //Set data range in the worksheet
                chart.DataRange = sheet.Range["A1:C6"];
                chart.IsSeriesInRows = false;

                IChartSerie chartSerie = chart.Series[0];

                //Set HasSeriesLines property to true.
                chartSerie.SerieFormat.CommonSerieOptions.HasSeriesLines = true;

                //Apply formats to SeriesLines.
                chartSerie.SerieFormat.CommonSerieOptions.PieSeriesLine.LineColor = Color.Green;

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
                workbook.SaveAs(Path.GetFullPath("Output/Chart.xlsx"));
                #endregion
            }
        }
    }
}





