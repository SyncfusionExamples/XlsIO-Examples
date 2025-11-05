using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Comments_to_PDF_at_End
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set print location to comments
                worksheet.PageSetup.PrintComments = ExcelPrintLocation.PrintSheetEnd;

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(worksheet);

                #region Save
                //Saving the workbook
                pdfDocument.Save(Path.GetFullPath("Output/CommentsToPDFAtEnd.pdf"));
                #endregion
            }
        }
    }
}





