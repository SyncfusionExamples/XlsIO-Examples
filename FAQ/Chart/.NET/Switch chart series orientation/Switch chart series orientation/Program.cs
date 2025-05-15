using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Switch_Chart_Series_Orientation
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

                //Add data for chart
                sheet.Range["A1"].Text = "Year";
                sheet.Range["B1"].Text = "2022";
                sheet.Range["C1"].Text = "2023";
                sheet.Range["D1"].Text = "2024";

                sheet.Range["A2"].Text = "Sales";
                sheet.Range["B2"].Number = 1000;
                sheet.Range["C2"].Number = 1500;
                sheet.Range["D2"].Number = 1800;

                sheet.Range["A3"].Text = "Profit";
                sheet.Range["B3"].Number = 200;
                sheet.Range["C3"].Number = 300;
                sheet.Range["D3"].Number = 400;

                //Create a Chart
                IChartShape chart = sheet.Charts.Add();

                //Set chart type
                chart.ChartType = ExcelChartType.Bar_Clustered;

                //Set data range in the worksheet
                chart.DataRange = sheet.Range["A1:D3"];

                //Set series orientation from rows to columns
                chart.IsSeriesInRows = false;

                //Positioning the chart in the worksheet
                chart.TopRow = 9;
                chart.LeftColumn = 1;
                chart.BottomRow = 22;
                chart.RightColumn = 8;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}
