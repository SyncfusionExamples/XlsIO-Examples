using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Chart_to_PDF
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
                IWorksheet worksheet = workbook.Worksheets[0];

                IChart chart = worksheet.Charts[0];

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();                

                //Convert Excel document with charts into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(chart);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ChartToPDF.pdf", FileMode.Create, FileAccess.Write);
                pdfDocument.Save(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ChartToPDF.pdf")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
