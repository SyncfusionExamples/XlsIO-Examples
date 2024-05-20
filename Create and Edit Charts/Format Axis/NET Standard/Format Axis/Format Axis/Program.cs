using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing;

namespace Format_Axis
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

                //Set the horizontal (category) axis title.
                chart.PrimaryCategoryAxis.Title = "Months";
                //Set the Vertical (value) axis title.
                chart.PrimaryValueAxis.Title = "Precipitation,in.";
                //Set title for secondary value axis
                chart.SecondaryValueAxis.Title = "Temperature,deg.F";

                //Customize the horizontal category axis.
                chart.PrimaryCategoryAxis.Border.LinePattern = ExcelChartLinePattern.Solid;
                chart.PrimaryCategoryAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryCategoryAxis.Border.LineWeight = ExcelChartLineWeight.Hairline;

                //Customize the vertical category axis.
                chart.PrimaryValueAxis.Border.LinePattern = ExcelChartLinePattern.Solid;
                chart.PrimaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.PrimaryValueAxis.Border.LineWeight = ExcelChartLineWeight.Narrow;

                //Customize the horizontal category axis font.
                chart.PrimaryCategoryAxis.Font.Color = ExcelKnownColors.Red;
                chart.PrimaryCategoryAxis.Font.FontName = "Calibri";
                chart.PrimaryCategoryAxis.Font.Bold = true;
                chart.PrimaryCategoryAxis.Font.Size = 8;

                //Customize the vertical category axis font.
                chart.PrimaryValueAxis.Font.Color = ExcelKnownColors.Red;
                chart.PrimaryValueAxis.Font.FontName = "Calibri";
                chart.PrimaryValueAxis.Font.Bold = true;
                chart.PrimaryValueAxis.Font.Size = 8;


                //Customize the secondary vertical category axis.
                chart.SecondaryValueAxis.Border.LinePattern = ExcelChartLinePattern.Solid;
                chart.SecondaryValueAxis.Border.LineColor = Syncfusion.Drawing.Color.Blue;
                chart.SecondaryValueAxis.Border.LineWeight = ExcelChartLineWeight.Narrow;

                //Customize the secondary vertical category axis font.
                chart.SecondaryValueAxis.Font.Color = ExcelKnownColors.Red;
                chart.SecondaryValueAxis.Font.FontName = "Calibri";
                chart.SecondaryValueAxis.Font.Bold = true;
                chart.SecondaryValueAxis.Font.Size = 8;

                //Axis title area text angle rotation.
                chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 270;

                //Maximum value in the axis.
                chart.PrimaryValueAxis.MaximumValue = 15;
                chart.PrimaryValueAxis.MinimumValue = 0;
                //Number format for axis.
                chart.PrimaryValueAxis.NumberFormat = "0.0";

                //Hiding major gridlines.
                chart.PrimaryValueAxis.HasMajorGridLines = true;

                //Showing minor gridlines.
                chart.PrimaryValueAxis.HasMinorGridLines = false;

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