using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System.IO;

namespace ConvertExcelToPDF
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load existing Excel file
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Sample.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Convert to PDF
                XlsIORenderer renderer = new XlsIORenderer();
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Create the MemoryStream to save the converted PDF   
                MemoryStream pdfStream = new MemoryStream();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Sample.pdf"), FileMode.Create, FileAccess.Write);
                pdfDocument.Save(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}