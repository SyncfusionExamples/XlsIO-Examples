using System.IO;
using Syncfusion.XlsIO;
using static System.Net.Mime.MediaTypeNames;

namespace Chart_Elements_ManualLayout
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

                //Manually resizing the data labels
                chart.Series[0].DataPoints[0].DataLabels.Layout.ManualLayout.Left = 0.09;
                chart.Series[0].DataPoints[2].DataLabels.Layout.ManualLayout.Top = 0;

                //Manually resizing the chart title area
                chart.ChartTitleArea.Text = "Sample Chart";
                chart.ChartTitleArea.Layout.ManualLayout.Top = 0.03;
                chart.ChartTitleArea.Layout.ManualLayout.Left = 0.02;

                //Manually positioning plot area
                chart.PlotArea.Layout.ManualLayout.LayoutTarget = LayoutTargets.inner;
                chart.PlotArea.Layout.ManualLayout.LeftMode = LayoutModes.edge;
                chart.PlotArea.Layout.ManualLayout.TopMode = LayoutModes.edge;

                //Manually resizing chart plot area
                chart.PlotArea.Layout.ManualLayout.Height = 0.59;
                chart.PlotArea.Layout.ManualLayout.Width = 0.81;
                chart.PlotArea.Layout.ManualLayout.Top = 0.18;
                chart.PlotArea.Layout.ManualLayout.Left = 0.16;

                //Manually positioning legend area
                chart.Legend.Layout.ManualLayout.LeftMode = LayoutModes.edge;
                chart.Legend.Layout.ManualLayout.TopMode = LayoutModes.edge;

                //Manually resizing chart legend area
                chart.Legend.Layout.ManualLayout.Height = 0.07;
                chart.Legend.Layout.ManualLayout.Width = 0.30;
                chart.Legend.Layout.ManualLayout.Top = 0.87;
                chart.Legend.Layout.ManualLayout.Left = 0.35;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ManualLayoutChart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ManualLayoutChart.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}