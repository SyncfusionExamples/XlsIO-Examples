using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Multiple_Excel_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream1 = new FileStream("../../../Data/Template1.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook1 = application.Workbooks.Open(inputStream1);

                FileStream inputStream2 = new FileStream("../../../Data/Template2.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook2 = application.Workbooks.Open(inputStream2);

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the first Excel document to PDF
                PdfDocument document = renderer.ConvertToPDF(workbook1);

                //Initialize XlsIORendererSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set the document as TemplateDocument
                settings.TemplateDocument = document;

                //Convert the second Excel document to PDF with renderer settings
                PdfDocument newDocument = renderer.ConvertToPDF(workbook2, settings);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("MultipleExcelToPDF.pdf", FileMode.Create, FileAccess.Write);
                newDocument.Save(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream1.Dispose();
                inputStream2.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("MultipleExcelToPDF.pdf")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
