using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Paper_Size
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"../../../Data/InputTemplate.xlsx"));

                //Set the paper size to A4 for all worksheets
                foreach (IWorksheet worksheet in workbook.Worksheets)
                {
                    worksheet.PageSetup.PaperSize = ExcelPaperSize.PaperA4;
                }

                //Initialize XlsIORendererSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set the layout option as FitAllColumnsOnOnePage
                settings.LayoutOptions = LayoutOptions.FitAllColumnsOnOnePage;

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the Excel document to PDF with renderer settings
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                //Save the workbook as PDF
                pdfDocument.Save(Path.GetFullPath("../../../Output/Output.pdf"));
            }
        }
    }
}