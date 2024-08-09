using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Picture_Hyperlink_in_Chart
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

                //Adding chart in the workbook
                IChart chart = workbook.Charts.Add();
                chart.DataRange = worksheet.Range["A1:C6"];
                chart.ChartType = ExcelChartType.Column_Clustered;
                chart.IsSeriesInRows = false;

                //Getting an image from the stream
                FileStream imageStream = new FileStream("../../../Data/Image.png", FileMode.Open, FileAccess.Read);
                Image image = Image.FromStream(imageStream);

                //Adding picture on the chart
                chart.Pictures.AddPicture(1, 1, imageStream);

                //Adding hyperlink to the picture on chart
                worksheet.HyperLinks.Add(workbook.Charts[0].Pictures[0] as IShape, ExcelHyperLinkType.Url, "https://www.Syncfusion.com", "click here");

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Chart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
                imageStream.Dispose();

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
