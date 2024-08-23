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
                FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                //Initialize fallBack font
                application.FallbackFonts.InitializeDefault();

                //Adds fallback font for "Arabic" script type.
                application.FallbackFonts.Add(ScriptType.Arabic, "Arial, Times New Roman");
                //Adds fallback font for "Hebrew" script type.
                application.FallbackFonts.Add(ScriptType.Hebrew, "Arial, Courier New");
                //Adds fallback font for "Thai" script type.
                application.FallbackFonts.Add(ScriptType.Thai, "Tahoma, Microsoft Sans Serif");
                //Adds fallback font for "Korean" script type.
                application.FallbackFonts.Add(ScriptType.Korean, "Malgun Gothic, Batang");

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
