using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using static System.Net.Mime.MediaTypeNames;
using System.Drawing;

namespace Format_Data_Labels
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

                for (int i = 0; i < chart.Series.Count; i++)
                {
                    //Enable the datalabel in chart.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                    //Set the data labels formatting - size, color, font name, position.
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Size = 10;
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Color = ExcelKnownColors.Black;
                    chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.FontName = "calibri";
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Bold = true;
                    chart.Series[i].DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;
                }

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