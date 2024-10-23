using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace Multithreading
{
    class MultiThreading
    {
        //Defines the number of threads to be created
        private const int ThreadCount = 1000;
        public static void Main()
        {
            //Create an array of threads based on the ThreadCount
            Thread[] threads = new Thread[ThreadCount];
            for (int i = 0; i < ThreadCount; i++)
            {
                threads[i] = new Thread(ReadEditConvertExcel);
                threads[i].Start();
            }

            //Ensure all threads complete by calling Join on each thread
            for (int i = 0; i < ThreadCount; i++)
            {
                threads[i].Join();
            }
        }

        //Method to convert Excel to PDF
        static void ReadEditConvertExcel()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                inputStream.Close();
                IWorksheet sheet = workbook.Worksheets[0];

                //Add text, formula, and number in the worksheet
                sheet.Range["A1"].Text = "Hello World" + DateTime.Now;
                Console.WriteLine(sheet.Range["A1"].Text);
                sheet.Range["A2"].Formula = "=Now()";
                sheet.Range["A3"].Number = 12345;

                //Convert the Excel workbook to PDF
                XlsIORenderer xlsIORenderer = new XlsIORenderer();
                PdfDocument pdfDocument = xlsIORenderer.ConvertToPDF(workbook);

                //Save the PDF document
                MemoryStream fileStream = new MemoryStream();
                pdfDocument.Save(fileStream);
                fileStream.Close();
                pdfDocument.Dispose();
            }
        }
    }
}