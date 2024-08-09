using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Chart_Series_Border
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding chart in the worksheet
                IChartShape chart = worksheet.Charts.Add();
                chart.DataRange = worksheet.Range["A1:C5"];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.IsSeriesInRows = false;

                //Accessing first chart series
                IChartSerie serie = chart.Series[0];

                //Formatting the series border
                serie.SerieFormat.LineProperties.LineColor = Color.Brown;
                serie.SerieFormat.LineProperties.LinePattern = ExcelChartLinePattern.CircleDot;
                serie.SerieFormat.LineProperties.LineWeight = ExcelChartLineWeight.Wide;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Chart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

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
