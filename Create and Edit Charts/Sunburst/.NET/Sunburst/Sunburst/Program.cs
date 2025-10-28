using System.IO;
using Syncfusion.XlsIO;

namespace Sunburst
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

                //Set chart type as Sunburst
                chart.ChartType = ExcelChartType.SunBurst;

                //Set data range in the worksheet
                chart.DataRange = sheet["A1:D16"];

                //Set the chart title
                chart.ChartTitle = "Sales by annual";

                //Formatting data labels      
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

                //Hiding the legend
                chart.HasLegend = false;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Sunburst.xlsx"));
                #endregion
            }
        }
    }
}





