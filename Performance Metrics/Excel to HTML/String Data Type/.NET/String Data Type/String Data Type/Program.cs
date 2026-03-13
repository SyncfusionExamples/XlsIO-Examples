using Syncfusion.XlsIO;

namespace String_Data_Type
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load existing Excel file
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

                //Create stream to store HTML file.
                Stream stream = new MemoryStream();

                //Save a workbook as HTML file
                workbook.SaveAsHtml(stream, Syncfusion.XlsIO.Implementation.HtmlSaveOptions.Default);

                stream.Dispose();
                workbook.Close();
            }
        }
    }
}