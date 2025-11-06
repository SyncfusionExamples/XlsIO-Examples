using Syncfusion.XlsIO;

namespace Excel_to_TSV
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

                //Save the workbook in CSV format with tab(\t) as delimiter
                workbook.SaveAs(Path.GetFullPath("Output/Excel to TSV.tsv"), "\t");
            }
        }
    }
}