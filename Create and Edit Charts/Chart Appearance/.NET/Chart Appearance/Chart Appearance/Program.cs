using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Chart_Appearance
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                IChartShape chart = sheet.Charts.Add();
                chart.DataRange = sheet.UsedRange;

                //Format Chart Area
                IChartFrameFormat chartArea = chart.ChartArea;
                //Fill Effects
                chartArea.Fill.FillType = ExcelFillType.Gradient;

                //Set chart area fill color
                chartArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Color.WhiteSmoke;

                //Format Plot Area
                IChartFrameFormat chartPlotArea = chart.PlotArea;
                //Fill Effects
                chartPlotArea.Fill.FillType = ExcelFillType.Gradient;

                //Set plot area fill color 
                chartPlotArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartPlotArea.Fill.ForeColor = Color.YellowGreen;

                //Positioning the chart in the worksheet
                chart.TopRow = 8;
                chart.LeftColumn = 1;
                chart.BottomRow = 23;
                chart.RightColumn = 8;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Chart.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





