﻿using Syncfusion.XlsIO;
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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
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

                //Set the background color
                chart.Legend.FrameFormat.Fill.ForeColorIndex = ExcelKnownColors.Yellow;

                //Set the font
                chart.Legend.TextArea.Bold = true;
                chart.Legend.TextArea.FontName = "Times New Roman";
                chart.Legend.TextArea.Size = 10;
                chart.Legend.TextArea.Strikethrough = false;

                //Remove the legend
                chart.Legend.LegendEntries[0].Delete();

                //Set Legend without overlapping the chart
                chart.Legend.IncludeInLayout = true;

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}




