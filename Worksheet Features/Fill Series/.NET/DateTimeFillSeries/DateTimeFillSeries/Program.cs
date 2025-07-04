using Syncfusion.XlsIO;

namespace DateTimeFillSeries
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

                //Assign datetime value to the cell
                worksheet["A1"].DateTime = new DateTime(2025, 1, 1);

                //Define the range
                IRange range = worksheet["A1:A50"];

                //Fill series using the years option, step value and stop value
                range.FillSeries(ExcelSeriesBy.Columns, ExcelFillSeries.Years, 2, new DateTime(2100, 1, 1));

                //Saving the workbook 
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}