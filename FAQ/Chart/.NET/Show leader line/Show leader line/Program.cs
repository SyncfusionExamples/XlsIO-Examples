using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Show_Leader_Line
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Add data
                sheet.Range["A1"].Text = "Fruit";
                sheet.Range["B1"].Text = "Quantity";
                sheet.Range["A2"].Text = "Apple";
                sheet.Range["A3"].Text = "Banana";
                sheet.Range["A4"].Text = "Cherry";
                sheet.Range["B2"].Number = 40;
                sheet.Range["B3"].Number = 30;
                sheet.Range["B4"].Number = 30;

                //Add a Pie chart 
                IChart chart = sheet.Charts.Add();
                chart.ChartType = ExcelChartType.Pie;
                chart.DataRange = sheet.Range["A1:B4"];
                chart.IsSeriesInRows = false;
                chart.ChartTitle = "Fruit Distribution";

                //Enable data labels with values, and leader lines
                IChartSerie series = chart.Series[0];
                series.DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                series.DataPoints.DefaultDataPoint.DataLabels.ShowLeaderLines = true;

                //Manually resizing data label area using Manual Layout
                chart.Series[0].DataPoints[0].DataLabels.Layout.ManualLayout.Left = 0.09;
                chart.Series[0].DataPoints[0].DataLabels.Layout.ManualLayout.Top = 0.01;

                #region Save
                //Saving the workbook
                workbook.SaveAs("Output.xlsx");
                #endregion
            }
        }
    }
}