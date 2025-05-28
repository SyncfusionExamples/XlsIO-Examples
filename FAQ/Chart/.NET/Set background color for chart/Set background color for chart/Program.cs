using System;
using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Set_Background_Color_For_Chart
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the first chart in the worksheet
                IChartShape chart = worksheet.Charts[0];

                //Applying background color for plot area
                chart.PlotArea.Fill.ForeColor = Color.LightYellow;

                //Applying background color for chart area
                chart.ChartArea.Fill.ForeColor = Color.LightGreen;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }

        }

    }
}
