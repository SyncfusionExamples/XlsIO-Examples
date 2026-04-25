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

            // Remove all ListObjects from the sheet
            // Iterate in reverse order to avoid index shifting issues
            for (int i = worksheet.ListObjects.Count - 1; i >= 0; i--)
            {
                IListObject listObject = worksheet.ListObjects[i];
                worksheet.ListObjects.Remove(listObject);
            }

            workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
        }
    }
}