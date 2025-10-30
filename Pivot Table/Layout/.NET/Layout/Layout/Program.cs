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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/PivotTable.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[1];

                IPivotTable pivotTable = worksheet.PivotTables[0];
                //Layout the pivot table.
                pivotTable.Layout();

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath(@"Output/PivotTable_Layout.xlsx"));
            }
        }
    }
}




