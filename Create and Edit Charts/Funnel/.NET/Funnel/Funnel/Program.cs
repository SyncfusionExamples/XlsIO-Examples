using System.IO;
using Syncfusion.XlsIO;

namespace Funnel
{
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

                //Create a chart
                IChartShape chart = sheet.Charts.Add();

                //Set chart type as Funnel
                chart.ChartType = ExcelChartType.Funnel;

                //Set data range in the worksheet
                chart.DataRange = sheet.Range["A1:B6"];

                //Set the chart title
                chart.ChartTitle = "Funnel";

                //Formatting the legend and data label option
                chart.HasLegend = false;
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Funnel.xlsx"));
                #endregion
            }
        }
    }
}





