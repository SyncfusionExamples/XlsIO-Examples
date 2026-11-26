using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            // Instantiate the Excel application object
            IApplication application = excelEngine.Excel;

            // Assign default application version
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open a new workbook contains table
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data\InputTemplate.xlsx"));

            // Access first worksheet from the workbook
            IWorksheet worksheet = workbook.Worksheets[0];

            IWorksheet pivotSheet = workbook.Worksheets[1];

            // Create pivot cache from the table location
            IPivotCache cache = workbook.PivotCaches.Add(worksheet.ListObjects[0].Location);

            IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);

            // Save the workbook to disk in XLSX format
            workbook.SaveAs(Path.GetFullPath(@"Output\Output.xlsx"));
        }
    }
}