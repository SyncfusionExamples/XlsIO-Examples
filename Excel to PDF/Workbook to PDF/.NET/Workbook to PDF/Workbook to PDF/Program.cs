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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                #region Save
                //Saving the workbook
                pdfDocument.Save(Path.GetFullPath("Output/WorkbookToPDF.pdf"));
                #endregion
            }
        }
    }
}