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

                //Adds fallback font for "Arabic" specific unicode range.
                application.FallbackFonts.Add(new FallbackFont(0x0600,0x06ff,"Arial"));
                //Adds fallback font for "Hebrew" specific unicode range.
                application.FallbackFonts.Add(new FallbackFont(0x0590, 0x05ff, "Times New Roman"));
                //Adds fallback font for "Thai" specific unicode range.
                application.FallbackFonts.Add(new FallbackFont(0x0E00, 0x0E7F, "Tahoma"));
                //Adds fallback font for "Korean" specific unicode range.
                application.FallbackFonts.Add(new FallbackFont(0xAC00, 0xD7A3, "Malgun Gothic"));

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Excel to PDF
                Stream stream = new FileStream("Sample.pdf", FileMode.Create, FileAccess.ReadWrite);
                pdfDocument.Save(stream);

                //Close and Dispose
                workbook.Close();
                stream.Dispose();
            }
        }
    }
}