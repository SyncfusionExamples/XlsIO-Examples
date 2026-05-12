using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        // Create a new Excel application instance
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath("Data/InputTemplate.xlsx"));

            foreach (IWorksheet worksheetxls in workbook.Worksheets)
            {
                worksheetxls.EnableSheetCalculations();
                worksheetxls.CalcEngine.AllowShortCircuitIFs = true;
                worksheetxls.CalcEngine.UseFormulaValues = true;
                // Enable circular reference handling
                worksheetxls.CalcEngine.ThrowCircularException = true;
                worksheetxls.CalcEngine.IterationMaxCount = 1000;
            }

            IWorksheet worksheet = workbook.Worksheets.First();
            Console.WriteLine(worksheet["I234"].CalculatedValue);
            Console.WriteLine(worksheet["J234"].CalculatedValue);

            workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
        }
    }
}
