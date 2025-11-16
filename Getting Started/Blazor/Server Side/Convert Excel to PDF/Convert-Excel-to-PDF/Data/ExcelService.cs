using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using Microsoft.AspNetCore.Components.RenderTree;

namespace Convert_Excel_to_PDF.Data
{
    public class ExcelService
    {
        public MemoryStream ConvertExceltoPDF()
        {
            // Open an existing Excel document.
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                // Open the workbook.
                IWorkbook workbook = application.Workbooks.Open(@"wwwroot/InputTemplate.xlsx");

                // Instantiate the Excel to PDF renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Create the MemoryStream to save the converted PDF.      
                MemoryStream pdfStream = new MemoryStream();

                //Save the converted PDF document to MemoryStream.
                pdfDocument.Save(pdfStream);
                pdfStream.Position = 0;

                return pdfStream;           
            }
        }
    }
}
