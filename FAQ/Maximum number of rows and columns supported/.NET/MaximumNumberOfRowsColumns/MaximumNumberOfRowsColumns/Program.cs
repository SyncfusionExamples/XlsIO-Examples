using Syncfusion.XlsIO;

namespace MaximumNumberOfRowsColumns
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //For getting maximum supported rows and columns  
                int maxRow = workbook.MaxRowCount;
                int maxColumns = workbook.MaxColumnCount;

                //Display maximum number of rows and columns supported
                Console.WriteLine("Maximum number of rows supported: " + maxRow.ToString());
                Console.WriteLine("Maximum number of columns supported: " + maxColumns.ToString());
            }
        }
        }
}