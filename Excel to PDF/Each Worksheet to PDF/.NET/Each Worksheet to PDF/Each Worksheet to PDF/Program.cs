using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;
using Syncfusion.XlsIORenderer;

namespace Each_Worksheet_to_PDF
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();
                PdfDocument pdfDocument = new PdfDocument();

                foreach (IWorksheet sheet in workbook.Worksheets)
                {
                    pdfDocument = renderer.ConvertToPDF(sheet);

                    #region Save
                    //Saving the workbook
                    FileStream outputStream = new FileStream(sheet.Name +".pdf", FileMode.Create, FileAccess.Write);
                    pdfDocument.Save(outputStream);
                    #endregion

                    //Dispose streams
                    outputStream.Dispose();
                    inputStream.Dispose();

                    System.Diagnostics.Process process = new System.Diagnostics.Process();
                    process.StartInfo = new System.Diagnostics.ProcessStartInfo(sheet.Name + ".pdf")
                    {
                        UseShellExecute = true
                    };
                    process.Start();
                }
            }
        }
    }
}

