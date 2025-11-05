using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Fit_All_Rows_On_One_Page
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

                //Initialize XlsIORendererSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set layout option as FitAllRowsOnOnePage
                settings.LayoutOptions = LayoutOptions.FitAllRowsOnOnePage;

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the Excel document to PDF with renderer settings
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                #region Save
                //Saving the workbook
                pdfDocument.Save(Path.GetFullPath("Output/FitAllRowsOnOnePage.pdf"));
                #endregion
            }
        }
    }
}