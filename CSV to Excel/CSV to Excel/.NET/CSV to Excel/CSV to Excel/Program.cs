using Syncfusion.XlsIO;

namespace CSV_to_Excel
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Open the CSV file
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.csv"), ",");
                IWorksheet worksheet = workbook.Worksheets[0];

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/CSV to Excel.xlsx"));
            }
        }
    }
}