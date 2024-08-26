using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing;

namespace Axis
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
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Set the axis title
                chart.PrimaryCategoryAxis.Title = "Months";
                chart.PrimaryValueAxis.Title = "Precipitation,in.";
                chart.SecondaryValueAxis.Title = "Temperature,deg.F";

                //Set the border 
                chart.PrimaryCategoryAxis.Border.LinePattern = ExcelChartLinePattern.CircleDot;
                chart.PrimaryCategoryAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryCategoryAxis.Border.LineWeight = ExcelChartLineWeight.Hairline;

                chart.PrimaryValueAxis.Border.LinePattern = ExcelChartLinePattern.CircleDot;
                chart.PrimaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryValueAxis.Border.LineWeight = ExcelChartLineWeight.Hairline;

                chart.SecondaryValueAxis.Border.LinePattern = ExcelChartLinePattern.Solid;
                chart.SecondaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.SecondaryValueAxis.Border.LineWeight = ExcelChartLineWeight.Hairline;

                //Set the font
                chart.PrimaryCategoryAxis.Font.Color = ExcelKnownColors.Red;
                chart.PrimaryCategoryAxis.Font.FontName = "Calibri";
                chart.PrimaryCategoryAxis.Font.Bold = true;
                chart.PrimaryCategoryAxis.Font.Size = 8;

                chart.PrimaryValueAxis.Font.Color = ExcelKnownColors.Red;
                chart.PrimaryValueAxis.Font.FontName = "Calibri";
                chart.PrimaryValueAxis.Font.Bold = true;
                chart.PrimaryValueAxis.Font.Size = 8;

                chart.SecondaryValueAxis.Font.Color = ExcelKnownColors.Red;
                chart.SecondaryValueAxis.Font.FontName = "Calibri";
                chart.SecondaryValueAxis.Font.Bold = true;
                chart.SecondaryValueAxis.Font.Size = 8;

                //Set the rotation
                chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 270;
                chart.SecondaryValueAxis.TitleArea.TextRotationAngle = 90;

                //Set the number format
                chart.PrimaryValueAxis.NumberFormat = "0.0";
                chart.SecondaryValueAxis.NumberFormat = "0.0";

                //Set maximum value
                chart.PrimaryValueAxis.MaximumValue = 14.0;
                chart.SecondaryValueAxis.MaximumValue = 49.5;

                //Set minimum value
                chart.PrimaryValueAxis.MinimumValue = 0;
                chart.SecondaryValueAxis.MinimumValue = 46.5;

                //Set maxcross
                chart.SecondaryValueAxis.IsMaxCross = true;

                //Set major tick mark
                chart.PrimaryCategoryAxis.MajorTickMark = ExcelTickMark.TickMark_Inside;
                chart.PrimaryValueAxis.MajorTickMark = ExcelTickMark.TickMark_Outside;
                chart.SecondaryValueAxis.MajorTickMark = ExcelTickMark.TickMark_Outside;

                //Showing major gridlines
                chart.PrimaryValueAxis.HasMajorGridLines = true;
                
                //Hiding minor gridlines
                chart.PrimaryValueAxis.HasMinorGridLines = false;

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




