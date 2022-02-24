using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace Workbook_to_PDF
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

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("WorkbookToPDF.pdf", FileMode.Create, FileAccess.Write);
                pdfDocument.Save(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("WorkbookToPDF.pdf")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
