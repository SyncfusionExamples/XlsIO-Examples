using Syncfusion.XlsIO;

namespace Excel_to_CSV
{
    class program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Excel to CSV.csv"), ",");
            }
        }
    }
}