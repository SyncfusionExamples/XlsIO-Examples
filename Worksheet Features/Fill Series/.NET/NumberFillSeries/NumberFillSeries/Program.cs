using Syncfusion.XlsIO;

namespace NumberFillSeries
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

                //Assign value to the cell
                worksheet["A1"].Number = 1;

                //Define the range
                IRange range = worksheet["A1:A100"];

                //Fill series using the linear option, step value and stop value
                range.FillSeries(ExcelSeriesBy.Columns, ExcelFillSeries.Linear, 5, 1000);

                //Saving the workbook 
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}