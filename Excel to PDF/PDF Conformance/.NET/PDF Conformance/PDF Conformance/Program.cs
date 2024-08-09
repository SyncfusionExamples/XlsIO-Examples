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
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

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
                FileStream outputStream = new FileStream("PDFConformance.pdf", FileMode.Create, FileAccess.Write);
                pdfDocument.Save(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PDFConformance.pdf")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
