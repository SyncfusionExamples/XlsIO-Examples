using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Format_Chart_Area
{
    class program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);  
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Format Chart Area
                IChartFrameFormat chartArea = chart.ChartArea;

                //Set border line pattern, color, line weight.
                chartArea.Border.LinePattern = ExcelChartLinePattern.Solid;
                chartArea.Border.LineColor = Color.Blue;
                chartArea.Border.LineWeight = ExcelChartLineWeight.Hairline;

                //Set fill type and fill colors.
                chartArea.Fill.FillType = ExcelFillType.Gradient;
                chartArea.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chartArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Color.White;

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
