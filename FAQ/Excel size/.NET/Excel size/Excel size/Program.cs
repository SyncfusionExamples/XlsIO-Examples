using System;
using System.IO;
using Syncfusion.XlsIO;

namespace ExcelSize
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
                IWorksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["A1"].Text = "Sample Data";

                //Save to memory stream
                using (MemoryStream stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    //Compute file size in bytes
                    long sizeInBytes = stream.Length;
                    Console.WriteLine($"File size: {sizeInBytes} bytes");

                    //Convert to KB 
                    double sizeInKB = sizeInBytes / 1024.0;
                    Console.WriteLine($"File size: {sizeInKB:F2} KB");
                }
            }
        }

    }
}