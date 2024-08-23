using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace ExcelToPDF
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "../../../Data/Invoice.xlsx";
            Stream inputExcelData = File.OpenRead(filePath);
            Stream outputPDFData = ConvertExcelToPDF(inputExcelData);
            File.WriteAllBytes(Path.GetFullPath(@"Output/Invoice.pdf"), ((MemoryStream)outputPDFData).ToArray());
        }

        static Stream ConvertExcelToPDF(Stream inputExcelData)
        {
            MemoryStream pdfStream = new MemoryStream();

            //Instantiate the spreadsheet creation engine.
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object.
                IApplication application = excelEngine.Excel;

                //Set the default application version.
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the existing Excel file into IWorkbook.
                IWorkbook workbook = application.Workbooks.Open(inputExcelData);

                //Settings for Excel to PDF conversion
                XlsIORendererSettings settings = new XlsIORendererSettings();
                
                //Set the layout option to fit all columns on one page.
                settings.LayoutOptions = LayoutOptions.FitAllColumnsOnOnePage;

                //Initialize the XlsIORenderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Initialize the PDF document.
                PdfDocument pdfDocument = new PdfDocument();

                //Convert the Excel document to PDF.
                pdfDocument = renderer.ConvertToPDF(workbook, settings);

                //Save the PDF file.
                pdfDocument.Save(pdfStream);

                //Close the PDF document.
                pdfDocument.Close();

                //Close the workbook.
                workbook.Close();
            }

            return pdfStream;
        }
    }
}
