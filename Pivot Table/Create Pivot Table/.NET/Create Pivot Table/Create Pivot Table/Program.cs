using System.IO;
using Syncfusion.XlsIO;

namespace Create_Pivot_Table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/PivotData.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
                IWorksheet pivotSheet = workbook.Worksheets[1];

                //Create Pivot cache with the given data range
                IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:H50"]);

                //Create "PivotTable1" with the cache at the specified range
                IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);

                //Add Pivot table fields (Row and Column fields)
                pivotTable.Fields[2].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[6].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[3].Axis = PivotAxisTypes.Column;

                //Add data field
                IPivotField field = pivotTable.Fields[5];
                pivotTable.DataFields.Add(field, "Sum", PivotSubtotalTypes.Sum);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/PivotTable.xlsx"));
                #endregion
            }
        }
    }
}