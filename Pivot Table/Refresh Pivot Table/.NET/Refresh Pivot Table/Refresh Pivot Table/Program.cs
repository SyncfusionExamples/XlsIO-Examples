using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace Refresh_Pivot_Table
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
                IWorksheet worksheet = workbook.Worksheets[0];

                //Updating a new value in the pivot data
                worksheet.SetValue(2, 3, "250");

                //Accessing the pivot table 
                IPivotTable pivotTable = workbook.Worksheets[1].PivotTables[0];
                PivotTableImpl pivotTableImpl = pivotTable as PivotTableImpl;

                //Refreshing pivot cache to update the pivot table
                pivotTableImpl.Cache.IsRefreshOnLoad = true;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RefreshPivotTable.xlsx"));
                #endregion
            }
        }
    }
}





