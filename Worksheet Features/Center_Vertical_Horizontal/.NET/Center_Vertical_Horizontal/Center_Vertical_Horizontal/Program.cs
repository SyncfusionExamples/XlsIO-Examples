using Syncfusion.XlsIORenderer;
using Syncfusion.XlsIO;
using Syncfusion.Pdf;

namespace Center_Vertical_Horizontal
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                for (int i = 1; i <= 10; i++)
                {
                    for (int j = 1; j <= 5; j++)
                    {
                        sheet.Range[i, j].Text = sheet.Range[i, j].AddressLocal;
                    }
                }

                foreach (IWorksheet worksheet in workbook.Worksheets)
                {
                    worksheet.PageSetup.CenterHorizontally = true;
                    worksheet.PageSetup.CenterVertically = true;
                }

                XlsIORenderer renderer = new XlsIORenderer();

                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);
                pdfDocument.Save(Path.GetFullPath("Output/Output.pdf"));
            }
        }
    }
}