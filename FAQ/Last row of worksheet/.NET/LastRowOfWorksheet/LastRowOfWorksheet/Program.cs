using Syncfusion.XlsIO;

namespace LastRowOfWorksheet
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

                //Get the last row from the used range
                int lastrow = workbook.ActiveSheet.UsedRange.LastRow;

                //Display the last row
                Console.WriteLine("The last row in the used range is: " + lastrow);

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }
        }
    }
}