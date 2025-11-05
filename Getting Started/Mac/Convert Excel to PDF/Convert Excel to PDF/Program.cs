using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace Excel_to_PDF_Mac
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load existing Excel file
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Sample.xlsx"));

                //Convert to PDF
                XlsIORenderer renderer = new XlsIORenderer();
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                #region Save
                //Saving the workbook
                pdfDocument.Save(Path.GetFullPath("Output/Sample.pdf"));
                #endregion
            }
        }
    }
}