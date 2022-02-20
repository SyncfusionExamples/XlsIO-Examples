using System.IO;
using Syncfusion.XlsIO;

namespace DataTable_in_Chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assigning data in the worksheet
                worksheet.Range["A1"].Text = "Items";
                worksheet.Range["B1"].Text = "Amount(in $)";
                worksheet.Range["C1"].Text = "Count";

                worksheet.Range["A2"].Text = "Beverages";
                worksheet.Range["A3"].Text = "Condiments";
                worksheet.Range["A4"].Text = "Confections";
                worksheet.Range["A5"].Text = "Dairy Products";
                worksheet.Range["A6"].Text = "Grains / Cereals";

                worksheet.Range["B2"].Number = 2776;
                worksheet.Range["B3"].Number = 1077;
                worksheet.Range["B4"].Number = 2287;
                worksheet.Range["B5"].Number = 1368;
                worksheet.Range["B6"].Number = 3325;

                worksheet.Range["C2"].Number = 925;
                worksheet.Range["C3"].Number = 378;
                worksheet.Range["C4"].Number = 880;
                worksheet.Range["C5"].Number = 581;
                worksheet.Range["C6"].Number = 189;

                //Adding a chart to the worksheet
                IChartShape chart = worksheet.Charts.Add();
                chart.DataRange = worksheet.Range["A1:C6"];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.IsSeriesInRows = false;

                //Adding title to the chart
                chart.ChartTitle = "Chart with Data Table";

                //Adding data table to the chart
                chart.HasDataTable = true;

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
                FileStream outputStream = new FileStream("Chart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Chart.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
