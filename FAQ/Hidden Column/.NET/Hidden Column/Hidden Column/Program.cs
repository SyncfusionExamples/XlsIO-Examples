using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);

            // Use the concrete WorksheetImpl when you need access to implementation-specific members
            WorksheetImpl sheet = workbook.Worksheets[0] as WorksheetImpl;

            // Hide column 1
            sheet.ShowColumn(1, false);

            // Detect whether column 1 is hidden
            bool hidden = sheet.ColumnInformation[1] != null && sheet.ColumnInformation[1].IsHidden;

            Console.WriteLine($"Column 1 hidden: {hidden}");

            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}