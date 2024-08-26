using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Legend
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];
                IChartShape chart = worksheet.Charts[0];

                //Add the legend
                chart.HasLegend = true;

                //Set the position
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                //View legend horizontally
                chart.Legend.IsVerticalLegend = false;

                //Set the border
                chart.Legend.FrameFormat.Border.AutoFormat = false;
                chart.Legend.FrameFormat.Border.IsAutoLineColor = false;
                chart.Legend.FrameFormat.Border.LineColor = Color.Black;
                chart.Legend.FrameFormat.Border.LinePattern = ExcelChartLinePattern.DashDot;
                chart.Legend.FrameFormat.Border.LineWeight = ExcelChartLineWeight.Narrow;

                //Set the color
                chart.Legend.TextArea.Color = ExcelKnownColors.Pink;

                //Set the font
                chart.Legend.TextArea.Bold = true;
                chart.Legend.TextArea.FontName = "Times New Roman";
                chart.Legend.TextArea.Size = 10;
                chart.Legend.TextArea.Strikethrough = false;

                //Remove the legend
                chart.Legend.LegendEntries[0].IsDeleted = true;

                //Set Legend without overlapping the chart
                chart.Legend.IncludeInLayout = true;

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




