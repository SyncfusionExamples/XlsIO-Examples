using Syncfusion.XlsIO;
using System;
using System.IO;

namespace CsvToExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
           using(ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                
                FileStream inputStream = new FileStream(@"../../../Data/PurchasedItems.csv", FileMode.Open,FileAccess.ReadWrite);

                //Opening CSV document with Comma separator
                IWorkbook workbook = application.Workbooks.Open(inputStream, ",");

                IWorksheet worksheet = workbook.Worksheets[0];

                //Applying filters to the data
                IAutoFilters filters = worksheet.AutoFilters;
                filters.FilterRange = worksheet[1, 1, worksheet.UsedRange.LastRow, worksheet.UsedRange.LastColumn];

                IAutoFilter filter = filters[1];

                filter.AddTextFilter("Wednesday");

                //Saving the CSV data as Excel
                FileStream outputStream = new FileStream(@"PurchasedItems.xlsx", FileMode.Create, FileAccess.ReadWrite);

                workbook.SaveAs(outputStream);
            }
        }
    }
}
