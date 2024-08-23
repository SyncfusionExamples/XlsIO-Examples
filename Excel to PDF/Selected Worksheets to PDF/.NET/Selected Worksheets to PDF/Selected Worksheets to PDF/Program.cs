using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Selected_Worksheets_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Open an Excel document
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Get the first worksheet
                IWorksheet worksheet1 = workbook.Worksheets[0];

                //Initialize XlsIORenderer
                XlsIORenderer renderer = new XlsIORenderer();

                //Initailize PdfDocument and convert first worksheet to PDF
                PdfDocument document = renderer.ConvertToPDF(worksheet1);

                //Initailize ExcelToPdfConverterSettings
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set the PdfDocument to TemplateDocument in ExcelToPdfConverterSettings
                settings.TemplateDocument = document;

                //Get the third worksheet
                IWorksheet worksheet3 = workbook.Worksheets[2];

                //Initailize new PdfDocument and convert third worksheet to PDF with settings
                PdfDocument newDocument = renderer.ConvertToPDF(worksheet3, settings);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("SelectedSheetsToPDF.pdf", FileMode.Create, FileAccess.Write);
                newDocument.Save(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

