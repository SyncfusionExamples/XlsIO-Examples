using Syncfusion.XlsIO;

namespace ChartNameInWorksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet worksheet;
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    // Access the worksheet 
                    worksheet = workbook.Worksheets[i];

                    // Check if worksheet is empty
                    if (worksheet.UsedCells.Length == 0)
                        Console.WriteLine("The worksheet with name \""+ worksheet.Name + "\" is empty");
                }
            }
        }
    }
}