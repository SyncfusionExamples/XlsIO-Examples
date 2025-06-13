using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Convert_CSV_to_PDF
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("Data/Sample.csv", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet sheet = workbook.Worksheets[0];

                //Auto-fit all columns in the used range to prevent cropping
                sheet.UsedRange.AutofitColumns();

                //Initialize XlsIO renderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert CSV document into PDF document
                PdfDocument pdfDocument = renderer.ConvertToPDF(sheet);

                //Saving the PDF document
                FileStream outputStream = new FileStream("Output.pdf", FileMode.Create, FileAccess.ReadWrite);
                pdfDocument.Save(outputStream);

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}