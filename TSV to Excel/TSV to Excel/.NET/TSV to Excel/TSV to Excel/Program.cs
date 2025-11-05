using Syncfusion.XlsIO;

namespace TSV_to_Excel
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Open the TSV file
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.tsv"), "\t");

                //Save the workbook
                workbook.SaveAs(Path.GetFullPath("Output/TSV to Excel.xlsx"));
            }
        }
    }
}