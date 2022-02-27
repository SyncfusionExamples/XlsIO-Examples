using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace Fit_All_Columns_On_One_Page
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Initialize XlsIORendererSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set layout option as FitAllColumnsOnOnePage
                settings.LayoutOptions = LayoutOptions.FitAllColumnsOnOnePage;

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert the Excel document to PDF with renderer settings
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("FitAllColumnsOnOnePage.pdf", FileMode.Create, FileAccess.Write);
                pdfDocument.Save(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("FitAllColumnsOnOnePage.pdf")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
