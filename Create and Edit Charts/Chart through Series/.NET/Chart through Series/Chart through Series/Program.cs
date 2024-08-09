using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Chart_through_Series
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

                //Inserts the sample data for the chart
                sheet.Range["A1"].Text = "Month";
                sheet.Range["B1"].Text = "Product A";
                sheet.Range["C1"].Text = "Product B";

                //Months
                sheet.Range["A2"].Text = "Jan";
                sheet.Range["A3"].Text = "Feb";
                sheet.Range["A4"].Text = "Mar";
                sheet.Range["A5"].Text = "Apr";
                sheet.Range["A6"].Text = "May";

                //Create a random Data
                Random r = new Random();
                for (int i = 2; i <= 6; i++)
                {
                    for (int j = 2; j <= 3; j++)
                    {
                        sheet.Range[i, j].Number = r.Next(0, 500);
                    }
                }
                IChartShape chart = sheet.Charts.Add();

                //Set chart type
                chart.ChartType = ExcelChartType.Line;

                //Set Chart Title
                chart.ChartTitle = "Product Sales comparison";

                //Set first serie
                IChartSerie productA = chart.Series.Add("ProductA");
                productA.Values = sheet.Range["B2:B6"];
                productA.CategoryLabels = sheet.Range["A2:A6"];

                //Set second serie
                IChartSerie productB = chart.Series.Add("ProductB");
                productB.Values = sheet.Range["C2:C6"];
                productB.CategoryLabels = sheet.Range["A2:A6"];

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
