using Syncfusion.XlsIO;

namespace ChartNameInWorksheet
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

                //Get the chart name 
                string chartName = worksheet.Charts[0].Name;
                //Display the chart name 
                Console.WriteLine("The name of the chart is: " + chartName);

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }
        }
    }
}