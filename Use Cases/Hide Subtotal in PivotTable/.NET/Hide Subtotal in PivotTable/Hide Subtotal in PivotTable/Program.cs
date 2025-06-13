using Syncfusion.XlsIO;

namespace Hide_Subtotal_in_PivotTable
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                IWorksheet worksheet = workbook.Worksheets[0];

                IWorksheet pivotSheet = workbook.Worksheets[1];

                //Create Pivot cache with the given data range
                IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:C9"]);

                //Create PivotTable with the cache at the specified location
                IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);

                //Add Pivot table field
                IPivotField regionField = pivotTable.Fields["Region"];
                regionField.Axis = PivotAxisTypes.Row;

                //Hide subtotals
                regionField.Subtotals = PivotSubtotalTypes.None;

                //Add Pivot table field
                IPivotField categoryField = pivotTable.Fields["Category"];
                categoryField.Axis = PivotAxisTypes.Row;

                //Hide subtotals
                categoryField.Subtotals = PivotSubtotalTypes.None;

                //Add data field
                IPivotField dataField = pivotTable.Fields["Sales"];
                pivotTable.DataFields.Add(dataField, "Total Sales", PivotSubtotalTypes.Sum);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}