using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace PDF_Conformance
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Initialize XlsIO renderer settings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                // Set the conformance for PDF/A-1b conversion
                settings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B;

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                #region Save
                //Saving the workbook
                pdfDocument.Save(Path.GetFullPath("Output/PDFConformance.pdf"));
                #endregion
            }
        }
    }
}