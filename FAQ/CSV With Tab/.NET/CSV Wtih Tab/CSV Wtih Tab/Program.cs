using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        // Create an instance of ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.csv"), "\t");
            IWorksheet worksheet = workbook.Worksheets[0];

            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}
