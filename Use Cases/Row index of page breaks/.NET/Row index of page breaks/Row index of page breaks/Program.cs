using Syncfusion.XlsIO;

class Program
{
    public static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(inputStream);
            IWorksheet sheet = workbook.Worksheets[0];

            //Access all horizontal page breaks
            for (int i = 0; i < sheet.HPageBreaks.Count; i++)
            {
                int rowIndex = sheet.HPageBreaks[i].Location.Row;
                Console.WriteLine($"Page break {i + 1} is at row: {rowIndex}");
            }

            //Dispose streams
            inputStream.Dispose();
        }
    }
}