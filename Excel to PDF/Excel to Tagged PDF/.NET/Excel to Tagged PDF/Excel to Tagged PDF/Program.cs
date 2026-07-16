using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Excel_to_Tagged_PDF
{
    class Program
    {
        static void Main()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

                //Initialize XlsIORendererSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set AutoTag to true to create a tagged PDF
                settings.AutoTag = true;

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the Excel document to PDF with renderer settings
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                //Save the PDF document
                pdfDocument.Save(Path.GetFullPath(@"Output/Output.pdf"));
            }
        }
    }
}