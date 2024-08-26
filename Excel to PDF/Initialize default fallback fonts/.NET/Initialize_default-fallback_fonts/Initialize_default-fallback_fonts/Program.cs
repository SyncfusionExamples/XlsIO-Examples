using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;

namespace Initialize_default_fallback_fonts
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                //Initialize fallBack font
                application.FallbackFonts.InitializeDefault();

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Save the converted PDF document to stream.
                FileStream stream = new FileStream("Sample.pdf", FileMode.Create, FileAccess.ReadWrite);
                pdfDocument.Save(stream);

                //Close and Dispose
                workbook.Close();
                stream.Dispose();
            }
        }
    }
}




