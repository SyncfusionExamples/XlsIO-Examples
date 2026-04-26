using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        // Create an instance of ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            // Set the default version to Excel 2016
            excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2016;
            // Load the workbook from the specified path
            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
            // Get the first worksheet in the workbook
            IWorksheet worksheet = workbook.Worksheets[0];

            //Get data from merged area
            int row = 2, col = 5;

            IRange range = worksheet[row, col];

            string data = range.IsMerged ? worksheet[range.MergeArea.Row, range.MergeArea.Column].Value : range.Value;

            Console.WriteLine(data);
            // Save the workbook to a new file
            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}