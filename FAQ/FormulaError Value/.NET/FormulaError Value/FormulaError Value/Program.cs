using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
            IWorksheet worksheet = workbook.Worksheets[0];

            IRange usedRange = worksheet.UsedRange;
            int firstrow = usedRange.Row;
            int lastrow = usedRange.LastRow;
            int firstcol = usedRange.Column;
            int lastcol = usedRange.LastColumn;

            for (int row = firstrow; row <= lastrow; row++)
            {
                for (int col = firstcol; col <= lastcol; col++)
                {
                    if (worksheet[row, col] != null && worksheet[row, col].HasFormulaErrorValue)
                    {
                        Console.WriteLine($"Formula error value: {worksheet[row, col].FormulaErrorValue} in Address: {worksheet[row, col].AddressLocal}");
                    }
                }
            }

            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}