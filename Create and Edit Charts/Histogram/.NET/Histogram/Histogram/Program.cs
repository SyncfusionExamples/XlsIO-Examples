using System.IO;
using Syncfusion.XlsIO;

namespace Histogram
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

                //Set chart type as Histogram       
                chart.ChartType = ExcelChartType.Histogram;

                //Set data range in the worksheet   
                chart.DataRange = sheet["A1:A15"];

                //Category axis bin settings        
                chart.PrimaryCategoryAxis.BinWidth = 8;

                //Gap width settings
                chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 6;

                //Set the chart title and axis title
                chart.ChartTitle = "Height Data";
                chart.PrimaryValueAxis.Title = "Number of students";
                chart.PrimaryCategoryAxis.Title = "Height";

                //Hiding the legend
                chart.HasLegend = false;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Histogram.xlsx"));
                #endregion
            }
        }
    }
}





