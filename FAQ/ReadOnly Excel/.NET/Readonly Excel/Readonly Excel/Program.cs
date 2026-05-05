using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            // Instantiate the Excel application object
            IApplication application = excelEngine.Excel;

            // Assigns default application version
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Create a workbook with 1 worksheet
            IWorkbook workbook = application.Workbooks.Create(1);

            // Access first worksheet from the workbook
            IWorksheet worksheet = workbook.Worksheets[0];

            // Adding text to a cell
            worksheet.Range["A1"].Text = "Hello World";

            // Protect the worksheet
            worksheet.Protect("syncfusion", ExcelSheetProtection.All);

            // Save the workbook to disk in XLSX format
            workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
        }
    }
}
