using System.IO;
using Syncfusion.XlsIO;

namespace Explode_Pie_Chart
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding pie chart in the worksheet
                IChartShape chart = worksheet.Charts.Add();
                chart.DataRange = worksheet.Range["A3:B7"];
                chart.ChartType = ExcelChartType.Pie;
                chart.IsSeriesInRows = false;

                //Showing the values of data points
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                //Exploding the pie chart to 40%
                chart.Series[0].SerieFormat.Percent = 40;

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                //Positioning the chart in the worksheet
                chart.TopRow = 9;
                chart.LeftColumn = 1;
                chart.BottomRow = 22;
                chart.RightColumn = 8;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Chart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Chart.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
