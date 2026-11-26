using Syncfusion.XlsIO;
namespace Get_Row_Height_and_Column_Width_in_Pixels
{
    class Program
{
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing file
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get row height in pixels
                int rowheight = worksheet.GetRowHeightInPixels(1);

                //Get column width in pixels
                int columnwidth = worksheet.GetColumnWidthInPixels(1);

                Console.WriteLine($"Row Height: {rowheight}");
                Console.WriteLine($"Column Width: {columnwidth}");

                //Dispose stream
                inputStream.Dispose();
            }
        }
    }
}