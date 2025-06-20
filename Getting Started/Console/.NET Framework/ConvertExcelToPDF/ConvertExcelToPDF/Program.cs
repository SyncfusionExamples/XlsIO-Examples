using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using System.IO;

namespace ConvertExcelToPDF
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream excelStream = new FileStream("Sample.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                //Initialize ExcelToPdfConverter
                ExcelToPdfConverter converter = new ExcelToPdfConverter(workbook);

                //Initialize PDF document
                PdfDocument pdfDocument = new PdfDocument();

                //Convert Excel document to PDF document
                pdfDocument = converter.Convert();

                //Save the converted PDF document
                pdfDocument.Save("Sample.pdf");
            }
        }
    }
}
