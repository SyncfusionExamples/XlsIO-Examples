using Syncfusion.XlsIO;
using System;
using System.IO;

namespace ExcelToCSV
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                FileStream inputStream = new FileStream(@Path.GetFullPath(@"Data/PurchasedItems.xlsx"), FileMode.Open, FileAccess.ReadWrite);

                //Opening CSV document with Comma separator
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Saving the Excel data as CSV
                FileStream outputStream = new FileStream(@"PurchasedItems.csv", FileMode.Create, FileAccess.ReadWrite);


                workbook.SaveAs(outputStream,",");
            }
        }
    }
}






