using System.IO;
using Syncfusion.XlsIO;

namespace Font_Settings_in_Chart
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

                //Adding a chart in Excel worksheet
                IChartShape chart = worksheet.Charts.Add();
                chart.DataRange = worksheet.Range["A1:C5"];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.IsSeriesInRows = false;

                //Displaying the data label values of chart series
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;
                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;                

                //Setting font name, size and color for chart legend
                chart.Legend.TextArea.FontName = "Tahoma";
                chart.Legend.TextArea.Size = 20;
                chart.Legend.TextArea.Color = ExcelKnownColors.Red;

                //Setting font name, size and color for data labels of first series
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.FontName = "Tahoma";
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 14;
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Color = ExcelKnownColors.Green;

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
