using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Drawing;
using System.Drawing;

class Program
{
    static void Main(string[] args)
    {
        //Initialize Excel Engine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            //Load the existing Excel document
            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data\InputTemplate.xlsx"));

            Image headerImage = Image.FromStream(File.OpenRead(Path.GetFullPath(@"Data\Syncfusion.png")));
            foreach (IWorksheet sheet in workbook.Worksheets)
            {
                // IMPORTANT: put the image placeholder in the header/footer text
                sheet.PageSetup.CenterHeader = "&G";
                sheet.PageSetup.CenterFooter = "&G";

                // then assign the Image object
                sheet.PageSetup.CenterHeaderImage = headerImage;
                sheet.PageSetup.CenterFooterImage = headerImage;
            }

            XlsIORenderer renderer = new XlsIORenderer();
            XlsIORendererSettings rendererSettings = new XlsIORendererSettings();
            rendererSettings.HeaderFooterOption.ShowHeader = true;
            rendererSettings.HeaderFooterOption.ShowFooter = true;


            using (PdfDocument tempDoc = renderer.ConvertToPDF(workbook, rendererSettings))
            {
                tempDoc.Save(Path.GetFullPath(@"Output\ConvertedDocument.pdf"));
            }
        }
    }
}
