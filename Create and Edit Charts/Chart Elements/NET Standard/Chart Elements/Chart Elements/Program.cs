using System.IO;
using Syncfusion.XlsIO;

namespace Chart_Elements
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
                IWorksheet sheet = workbook.Worksheets[0];

                IChartShape chart = sheet.Charts[0];

                //Positioning chart in a worksheet
                chart.TopRow = 5;
                chart.LeftColumn = 5;
                chart.RightColumn = 10;
                chart.BottomRow = 10;

                //Manually positioning plot area
                chart.PlotArea.Layout.LayoutTarget = LayoutTargets.inner;
                chart.PlotArea.Layout.LeftMode = LayoutModes.edge;
                chart.PlotArea.Layout.TopMode = LayoutModes.edge;

                //Manually positioning chart legend
                chart.Legend.Layout.LeftMode = LayoutModes.edge;
                chart.Legend.Layout.TopMode = LayoutModes.edge;
                IShape chartShape = chart as IShape;

                //Set Height of the chart in pixels
                chartShape.Height = 300;

                //Set Width of the chart
                chartShape.Width = 500;

                //Manually resizing chart plot area
                chart.PlotArea.Layout.Left = 50;
                chart.PlotArea.Layout.Top = 75;
                chart.PlotArea.Layout.Width = 300;
                chart.PlotArea.Layout.Height = 200;

                //Manually resizing chart legend
                chart.Legend.Layout.Left = 400;
                chart.Legend.Layout.Top = 150;
                chart.Legend.Layout.Width = 200;
                chart.Legend.Layout.Height = 100;

                //Applying transparency to chart area
                chart.ChartArea.Fill.Transparency = 0.9;

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
