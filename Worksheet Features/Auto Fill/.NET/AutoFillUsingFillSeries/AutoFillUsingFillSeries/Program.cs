using Syncfusion.XlsIO;


namespace AutoFillUsingFillSeries
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

                //Define the source range
                IRange source = worksheet["A1:A3"];

                //Define the destination range
                IRange destinationRange = worksheet["A4:A100"];

                //Auto fill using the series option
                source.AutoFill(destinationRange, ExcelAutoFillType.FillSeries);

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}

