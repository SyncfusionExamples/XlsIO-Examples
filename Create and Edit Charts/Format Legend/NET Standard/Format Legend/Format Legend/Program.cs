using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing;

namespace Format_Legend
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Enable the legend.
                chart.HasLegend = true;

                //Set the position of legend.
                chart.Legend.Position = ExcelLegendPosition.Right;

                //Sets the legend border format - color, pattern, weight.
                chart.Legend.FrameFormat.Border.AutoFormat = false;
                chart.Legend.FrameFormat.Border.IsAutoLineColor = false;
                chart.Legend.FrameFormat.Border.LineColor = Syncfusion.Drawing.Color.Black;
                chart.Legend.FrameFormat.Border.LinePattern = ExcelChartLinePattern.DashDot;
                chart.Legend.FrameFormat.Border.LineWeight = ExcelChartLineWeight.Narrow;

                //Set the legend's text area format - font name, bold, color, size.
                chart.Legend.TextArea.Bold = true;
                chart.Legend.TextArea.Color = ExcelKnownColors.Pink;
                chart.Legend.TextArea.FontName = "Times New Roman";
                chart.Legend.TextArea.Size = 10;
                chart.Legend.TextArea.Strikethrough = false;

                //View legend in vertical.
                chart.Legend.IsVerticalLegend = true;

                //Modifies the legend entry.
                chart.Legend.LegendEntries[0].IsDeleted = true;

                //Manually resizing chart legend area using Layout.
                chart.Legend.Layout.Left = 0.2;
                chart.Legend.Layout.Top = 5;
                chart.Legend.Layout.Width = 60;
                chart.Legend.Layout.Height = 40;

                //Legend without overlapping the chart.
                chart.Legend.IncludeInLayout = true;

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}