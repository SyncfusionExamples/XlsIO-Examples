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
                FileStream excelStream = new FileStream("Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Create the FileStream to save the converted PDF.
                FileStream pdfStream = new FileStream("Output.pdf", FileMode.Create, FileAccess.ReadWrite);
                pdfDocument.Save(pdfStream);
            }
        }
    }
}




