using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace Gridlines_in_PDF
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

                //Set the gridlines display style as Invisible. Available options are Auto, Visible, Invisible
                settings.DisplayGridLines = GridLinesDisplayStyle.Visible;

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the Excel document to PDF with renderer settings
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                #region Save
                //Saving the workbook
                pdfDocument.Save(Path.GetFullPath("Output/Gridlines.pdf"));
                #endregion
            }
        }
    }
}





