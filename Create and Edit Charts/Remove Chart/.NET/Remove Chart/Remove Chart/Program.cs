using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Chart
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
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Remove the chart from the worksheet
                chart.Remove();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Chart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

