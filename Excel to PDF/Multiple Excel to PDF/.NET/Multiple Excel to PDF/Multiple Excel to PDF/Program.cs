using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Multiple_Excel_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook1 = application.Workbooks.Open(Path.GetFullPath(@"Data/Template1.xlsx"));

                IWorkbook workbook2 = application.Workbooks.Open(Path.GetFullPath(@"Data/Template2.xlsx"));

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the first Excel document to PDF
                PdfDocument document = renderer.ConvertToPDF(workbook1);

                //Initialize XlsIORendererSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set the document as TemplateDocument
                settings.TemplateDocument = document;

                //Convert the second Excel document to PDF with renderer settings
                PdfDocument newDocument = renderer.ConvertToPDF(workbook2, settings);

                #region Save
                //Saving the workbook
                newDocument.Save(Path.GetFullPath("Output/MultipleExcelToPDF.pdf"));
                #endregion
            }
        }
    }
}
