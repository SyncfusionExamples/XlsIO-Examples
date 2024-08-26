using System.IO;
using Syncfusion.XlsIO;

namespace Chart_from_Scratch
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

                object[] yValues = new object[] { 2000, 1000, 1000 };
                object[] xValues = new object[] { "Total Income", "Expenses", "Profit" };

                //Adding series and values
                IChartShape chart = sheet.Charts.Add();
                IChartSerie serie = chart.Series.Add(ExcelChartType.Pie);

                //Enters the X and Y values directly
                serie.EnteredDirectlyValues = yValues;
                serie.EnteredDirectlyCategoryLabels = xValues;

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                //Positioning the chart in the worksheet
                chart.TopRow = 1;
                chart.LeftColumn = 1;
                chart.BottomRow = 16;
                chart.RightColumn = 8;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Chart.xlsx"), FileMode.Create, FileAccess.Write);
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
