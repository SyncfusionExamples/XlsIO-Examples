﻿using System.IO;
using Syncfusion.XlsIO;

namespace Chart_Bars_Spacing
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding chart in the worksheet
                IChartShape chart = worksheet.Charts.Add();
                chart.DataRange = worksheet.Range["A1:C5"];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.IsSeriesInRows = false;

                //Adding space between bars of different series of single category
                chart.Series[0].SerieFormat.CommonSerieOptions.Overlap = 60;

                //Adding space between bars of different categories
                chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 80;

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





