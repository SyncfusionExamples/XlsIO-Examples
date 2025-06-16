using Syncfusion.XlsIO;

namespace FillSeriesByEnablingTrend
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign values to the cells
                worksheet["A1"].Number = 2;
                worksheet["A2"].Number = 4;
                worksheet["A3"].Number = 6;

                //Define the range
                IRange range = worksheet["A1:A100"];

                //Fill series using the linear option by enabling trend
                range.FillSeries(ExcelSeriesBy.Columns, ExcelFillSeries.Linear, true);

                //Saving the workbook 
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}