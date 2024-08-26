using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Drop_Lines
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine engine = new ExcelEngine())
            {
                IApplication application = engine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a Chart
                IChartShape chart = sheet.Charts.Add();

                //Set Chart Type
                chart.ChartType = ExcelChartType.Line;

                //Set data range in the worksheet
                chart.DataRange = sheet.Range["A1:C6"];
                chart.IsSeriesInRows = false;

                IChartSerie chartSerie = chart.Series[0];

                //Set HasDropLines property to true.
                chartSerie.SerieFormat.CommonSerieOptions.HasDropLines = true;

                //Apply formats to DropLines.
                chartSerie.SerieFormat.CommonSerieOptions.DropLines.LineColor = Color.Green;

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
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Chart.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

