using System.IO;
using Syncfusion.XlsIO;

namespace Waterfall
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a chart
                IChartShape chart = sheet.Charts.Add();

                //Set chart type as Waterfall
                chart.ChartType = ExcelChartType.WaterFall;

                //Set data range in the worksheet
                chart.DataRange = sheet["A2:B8"];

                //Data point settings as total in chart
                chart.Series[0].DataPoints[3].SetAsTotal = true;
                chart.Series[0].DataPoints[6].SetAsTotal = true;

                //Showing the connector lines between data points
                chart.Series[0].SerieFormat.ShowConnectorLines = true;

                //Set the chart title
                chart.ChartTitle = "Company Profit (in USD)";

                //Formatting data label and legend option
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;
                chart.Legend.Position = ExcelLegendPosition.Right;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Waterfall.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

