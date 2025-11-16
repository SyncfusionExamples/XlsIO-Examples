using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using System;
using System.IO;
using Syncfusion.XlsIORenderer;
using static System.Net.Mime.MediaTypeNames;

namespace Convert_Excel_to_PDF
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

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Save the converted PDF.
                pdfDocument.Save("Output.pdf");
            }
        }
    }
}