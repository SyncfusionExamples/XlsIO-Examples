using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Empty_Excel_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                IWorkbook workbook = application.Workbooks.Create(1);

                //Initialize XlsIORendererSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Enabling ThrowWhenExcelFileIsEmpty throws exception
                settings.ThrowWhenExcelFileIsEmpty = true;

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the Excel document to PDF with renderer settings
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                #region Save
                //Saving the workbook
                pdfDocument.Save(Path.GetFullPath("Output/EmptyExcelToPDF.pdf"));
                #endregion
            }
        }
    }
}