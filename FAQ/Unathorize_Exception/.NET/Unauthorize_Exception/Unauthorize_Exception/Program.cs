using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        //Create an instance of ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            //Set the default version as Excel 2016
            excelEngine.Excel.DefaultVersion = ExcelVersion.Excel2016;

            File.SetAttributes(Path.GetFullPath("Data/InputTemplate.xlsm"), FileAttributes.Normal);

            //Create a workbook
            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(Path.GetFullPath("Data/InputTemplate.xlsm"));

            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A2"].Text = "Hello, World!";

            File.SetAttributes(Path.GetFullPath("Data/InputTemplate.xlsm"), FileAttributes.Hidden);

            //Save the workbook to disk in xlsx format
            workbook.SaveAs(Path.GetFullPath("Output/Output.xlsm"));
        }
    }
}
