using System;
using Syncfusion.XlsIO;

namespace ChartProperties
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                #region Workbook Initialization
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
                #endregion

                //Create a Chart
                IChartShape chart = workbook.Worksheets[0].Charts.Add();

                chart.DataRange = worksheet.Range["A3:C15"];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.IsSeriesInRows = false;

                //Formatting the chart
                chart.ChartTitle = "Crescent City, CA";
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Size = 14;
                chart.ChartTitleArea.Bold = true;
                chart.ChartTitleArea.Color = ExcelKnownColors.Red;

                //Embedded Chart Position
                chart.TopRow = 2;
                chart.BottomRow = 30;
                chart.LeftColumn = 5;
                chart.RightColumn = 18;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}
