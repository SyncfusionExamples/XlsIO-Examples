using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace Layout
{
    class Prgoram
    {
        public static void Main(String[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/PivotTable.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[1];

                IPivotTable pivotTable = worksheet.PivotTables[0];
                //Layout the pivot table.
                pivotTable.Layout();

                string fileName = "PivotTable_Layout.xlsx";
                //Saving the workbook as stream
                FileStream stream = new FileStream(Path.GetFullPath(@"Output/PivotTable_Layout.xlsx"), FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }
    }
}




