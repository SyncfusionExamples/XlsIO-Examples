﻿using System.IO;
using Syncfusion.XlsIO;

namespace Box_and_Whisker
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
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a chart
                IChartShape chart = sheet.Charts.Add();

                //Set the chart title
                chart.ChartTitle = "Test Scores";

                //Set chart type as Box and Whisker
                chart.ChartType = ExcelChartType.BoxAndWhisker;

                //Set data range in the worksheet
                chart.DataRange = sheet["A1:D16"];

                //Box and Whisker settings on first series
                IChartSerie seriesA = chart.Series[0];
                seriesA.SerieFormat.ShowInnerPoints = false;
                seriesA.SerieFormat.ShowOutlierPoints = true;
                seriesA.SerieFormat.ShowMeanMarkers = true;
                seriesA.SerieFormat.ShowMeanLine = false;
                seriesA.SerieFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;

                //Box and Whisker settings on second series   
                IChartSerie seriesB = chart.Series[1];
                seriesB.SerieFormat.ShowInnerPoints = false;
                seriesB.SerieFormat.ShowOutlierPoints = true;
                seriesB.SerieFormat.ShowMeanMarkers = true;
                seriesB.SerieFormat.ShowMeanLine = false;
                seriesB.SerieFormat.QuartileCalculationType = ExcelQuartileCalculation.InclusiveMedian;

                //Box and Whisker settings on third series   
                IChartSerie seriesC = chart.Series[2];
                seriesC.SerieFormat.ShowInnerPoints = false;
                seriesC.SerieFormat.ShowOutlierPoints = true;
                seriesC.SerieFormat.ShowMeanMarkers = true;
                seriesC.SerieFormat.ShowMeanLine = false;
                seriesC.SerieFormat.QuartileCalculationType = ExcelQuartileCalculation.ExclusiveMedian;

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("BoxandWhisker.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("BoxandWhisker.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
