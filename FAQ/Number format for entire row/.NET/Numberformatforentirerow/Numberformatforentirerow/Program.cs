using Syncfusion.XlsIO;

class Program
{        
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];
            
            worksheet["A1"].Number = 1000.500;
            worksheet["B1"].Number = 1234;
            worksheet["C1"].Number = 54321.500;
            worksheet["D1"].Number = .500;

            worksheet["A1"].EntireRow.NumberFormat = "#,##0.0000";
            workbook.SaveAs("../../../Output/NumberFormats.xlsx");
        }
    }
}