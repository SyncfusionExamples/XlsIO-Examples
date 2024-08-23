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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Updating a new value in the pivot data
                worksheet.Range["C2"].Value = "250";

                //Accessing the pivot table 
                IPivotTable pivotTable = workbook.Worksheets[1].PivotTables[0];
                PivotTableImpl pivotTableImpl = pivotTable as PivotTableImpl;

                //Refreshing pivot cache to update the pivot table
                pivotTableImpl.Cache.IsRefreshOnLoad = true;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("RefreshPivotTable.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

