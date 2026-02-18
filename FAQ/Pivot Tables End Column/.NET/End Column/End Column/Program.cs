using Syncfusion.XlsIO;
using System;
using System.IO;

namespace End_Column
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
                IWorksheet sheet = workbook.Worksheets[1];
                IPivotTable pivotTable = sheet.PivotTables[0];

                // Ensure layout is calculated
                pivotTable.Layout();

                // Read EndLocation from the implementation type
                IRange endRange = (pivotTable as Syncfusion.XlsIO.Implementation.PivotTables.PivotTableImpl).EndLocation;
                int lastColumn = endRange.LastColumn;

                // Use lastColumn as needed (e.g., log)
                Console.WriteLine("PivotTable last column: " + lastColumn);
            }
        }
    }
}





