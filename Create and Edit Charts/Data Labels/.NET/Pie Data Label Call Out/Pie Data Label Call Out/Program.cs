using System;
using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Pie_Data_Label_CallOuts
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
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assigning data to cells
                worksheet.Range["A1"].Text = "Category";
                worksheet.Range["B1"].Text = "Value";
                worksheet.Range["A2"].Text = "Apples";
                worksheet.Range["B2"].Number = 30;
                worksheet.Range["A3"].Text = "Bananas";
                worksheet.Range["B3"].Number = 45;
                worksheet.Range["A4"].Text = "Cherries";
                worksheet.Range["B4"].Number = 25;

                //Add a pie chart to the worksheet
                IChartShape chart = worksheet.Charts.Add();

                //Set data range for the chart
                chart.DataRange = worksheet.Range["A1:B4"];

                //Specify chart type
                chart.ChartType = ExcelChartType.Pie;

                //Set chart properties
                chart.IsSeriesInRows = false;
                chart.ChartTitle = "Fruit Distribution";
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Right;

                //Position the chart within the worksheet
                chart.TopRow = 6;
                chart.LeftColumn = 1;
                chart.BottomRow = 20;
                chart.RightColumn = 10;

                //Customize data label for the first data point
                IChartSerie series = chart.Series[0];                      
                series.DataPoints[0].DataLabels.IsCategoryName = true;
                series.DataPoints[0].DataLabels.IsValue = true;

                //Enable data label callouts for the first data point
                series.DataPoints[0].DataLabels.ShowLeaderLines = true;

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