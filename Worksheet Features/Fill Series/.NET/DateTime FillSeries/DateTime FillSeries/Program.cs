using Syncfusion.XlsIO;

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

            //Setting the values to the cells
            worksheet["A1"].DateTime = new DateTime(2025, 1, 1);

            //Define the range
            IRange range = worksheet["A1:A10"];

            //Use FillSeries method to fill the values based on ExcelFillSeries
            range.FillSeries(ExcelSeriesBy.Columns, ExcelFillSeries.Years, 3, new DateTime(2060, 1, 1));

            //Saving the workbook 
            FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);

            //Dispose streams
            outputStream.Dispose();
        }

    }
}