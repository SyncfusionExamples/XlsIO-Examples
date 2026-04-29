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

            // Merge: true preserves top-left value and copies top-left formatting to merged area
            worksheet.Range["B8:C11"].Merge();

            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}
