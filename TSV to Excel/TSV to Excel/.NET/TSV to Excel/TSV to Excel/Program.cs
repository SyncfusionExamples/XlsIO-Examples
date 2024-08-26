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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.tsv"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, "\t");

                //Save the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




