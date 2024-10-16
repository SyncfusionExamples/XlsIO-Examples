using Syncfusion.XlsIO;
using System;

namespace GaugeChart
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Adding values in worksheet
                sheet.Range["A1"].Value = "Value";
                sheet.Range["A2"].Value = "30";
                sheet.Range["A3"].Value = "60";
                sheet.Range["A4"].Value = "90";
                sheet.Range["A5"].Value = "180";
                sheet.Range["C2"].Value = "value";
                sheet.Range["C3"].Value = "pointer";
                sheet.Range["C4"].Value = "End";
                sheet.Range["D2"].Value = "10";
                sheet.Range["D3"].Value = "1";
                sheet.Range["D4"].Value = "189";

                //Adding doughnut chart in worksheet
                IChartShape chart = sheet.Charts.Add();
                chart.ChartType = ExcelChartType.Doughnut;
                chart.DataRange = sheet.Range["A1:A5"];
                chart.IsSeriesInRows = false;

                //Formatting value series
                chart.Series["Value"].SerieFormat.CommonSerieOptions.DoughnutHoleSize = 60;
                chart.Series["Value"].SerieFormat.CommonSerieOptions.FirstSliceAngle = 270;
                chart.Series["Value"].DataPoints[3].DataFormat.Fill.Visible = false;

                //Adding pointer series as Pie chart
                chart.Series.Add("Pointer");
                chart.Series["Pointer"].SerieType = ExcelChartType.Pie;
                chart.Series["Pointer"].Values = sheet.Range["D2:D4"];
                chart.Series["Pointer"].UsePrimaryAxis = false;

                //Formatting pointer series
                chart.Series["Pointer"].SerieFormat.CommonSerieOptions.FirstSliceAngle = 270;
                chart.Series["Pointer"].DataPoints[0].DataFormat.Fill.Visible = false;
                chart.Series["Pointer"].DataPoints[1].DataFormat.Fill.ForeColorIndex = ExcelKnownColors.Black;
                chart.Series["Pointer"].DataPoints[2].DataFormat.Fill.Visible = false;

                //Saving the workbook as stream
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
            }
        }
    }
}