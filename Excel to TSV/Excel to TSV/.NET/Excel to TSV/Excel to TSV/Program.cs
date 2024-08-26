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

                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Save the workbook in CSV format with tab(\t) as delimiter
                FileStream outputStream = new FileStream("Output.tsv", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream, "\t");
            }
        }
    }
}




