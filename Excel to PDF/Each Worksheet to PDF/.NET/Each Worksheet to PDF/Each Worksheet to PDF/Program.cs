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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();
                PdfDocument pdfDocument = new PdfDocument();

                foreach (IWorksheet sheet in workbook.Worksheets)
                {
                    pdfDocument = renderer.ConvertToPDF(sheet);

                    #region Save
                    //Saving the workbook
                    pdfDocument.Save(sheet.Name +".pdf");
                    #endregion

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





