using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Set up sample data
            worksheet["A1"].Value2 = 10;
            worksheet["A2"].Value2 = 20;
            worksheet["A3"].Value2 = 30;

            // Create formulas
            worksheet["B1"].Formula = "=A1*2";
            worksheet["B2"].Formula = "=A2*2";
            worksheet["B3"].Formula = "=A3*2";

            worksheet.EnableSheetCalculations();
            // Move range B1:B3 to C1:C3
            worksheet["B1:B3"].MoveTo(worksheet["C1:C3"]);

            // Clear the formula info table to ensure dependent formulas recalculate
            worksheet.CalcEngine.FormulaInfoTable.Clear();

            // Now the formulas will recalculate correctly when their values are accessed
            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}
