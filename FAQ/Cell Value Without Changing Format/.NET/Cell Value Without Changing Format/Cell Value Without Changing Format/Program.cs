using Syncfusion.XlsIO;

class Program
{
    static void Main()
    {
        // Create an instance of ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Preserve existing formatting by assigning text directly
            worksheet.Range["A1"].Text = "1-";

            // Or set the cell's NumberFormat to Text before using Value
            worksheet.Range["A2"].NumberFormat = "@";
            worksheet.Range["A2"].Value = "1-";

            workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
        }
    }
}