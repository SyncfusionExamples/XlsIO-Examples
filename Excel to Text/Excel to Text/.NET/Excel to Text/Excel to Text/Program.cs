using Syncfusion.XlsIO;

namespace Excel_to_Text
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Initialize Excel engine and application.
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                // Open an existing workbook.
                FileStream inputStream = new FileStream(Path.GetFullPath("Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                // Save the workbook in .txt format with space (" ") as the delimiter.
                using FileStream outputStream = new FileStream(Path.GetFullPath("Output/Excel to Text.txt"), FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream, " ");
            }
        }
    }
}