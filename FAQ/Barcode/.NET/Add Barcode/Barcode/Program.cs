using System;
using System.IO;
using Syncfusion.XlsIO;

namespace AddBarcode
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

                // Load barcodes from local files
                FileStream barcode1 = new FileStream("Data/Barcode1.png", FileMode.Open, FileAccess.Read);
                FileStream barcode2 = new FileStream("Data/Barcode2.png", FileMode.Open, FileAccess.Read);
                worksheet.Pictures.AddPicture(1, 1, barcode1);
                worksheet.Pictures.AddPicture(15, 1, barcode2);
                worksheet.Pictures.AddPicture(1, 10, barcode1);
                worksheet.Pictures.AddPicture(15, 10, barcode2);

                // Save to file system
                FileStream stream = new FileStream("Output/Output.xlsx", FileMode.OpenOrCreate, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                workbook.Close();
                excelEngine.Dispose();
            }

        }

    }
}