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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

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

                //Save the PDF document 
                pdfDocument.Save("Sample.pdf");

                //Close and Dispose
                workbook.Close();
            }
        }
    }
}




