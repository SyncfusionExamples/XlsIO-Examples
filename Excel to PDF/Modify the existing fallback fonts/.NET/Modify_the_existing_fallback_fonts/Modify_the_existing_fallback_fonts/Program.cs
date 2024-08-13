using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using static System.Net.Mime.MediaTypeNames;
using Syncfusion.Office;

namespace Fallback_fonts_based_in_scripttype
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream fileStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                //Initialize fallBack font
                application.FallbackFonts.InitializeDefault();

                FallbackFonts fallbackFonts = application.FallbackFonts;
                foreach(FallbackFont fallbackFont in fallbackFonts)
                {
                    //Customize a default fallback font name as "David" for the Hebrew script.
                    if (fallbackFont.ScriptType == ScriptType.Hebrew)
                        fallbackFont.FontNames = "David";
                }

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Save the PDF document to stream
                FileStream stream = new FileStream("Sample.pdf", FileMode.Create, FileAccess.ReadWrite);
                pdfDocument.Save(stream);

                //Close and Dispose
                workbook.Close();
                stream.Dispose();
            }
        }
    }
}