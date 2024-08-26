using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace Expand_or_Collapse_Pivot_Rows
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
                IWorksheet pivotSheet = workbook.Worksheets[1];

                //Create pivot cache with the given data range
                IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:H50"]);

                //Create "PivotTable1" with the cache at the specified range
                IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);

                //Add pivot table fields (Row and Column fields)
                pivotTable.Fields[0].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[1].Axis = PivotAxisTypes.Row;

                //Add data field
                IPivotField field = pivotTable.Fields[2];
                pivotTable.DataFields.Add(field, "Sum", PivotSubtotalTypes.Sum);

                //Initialize PivotItemOptions
                PivotItemOptions options = new PivotItemOptions();
                options.IsHiddenDetails = false;

                //Collapsing the first and second items of the first pivot field using PivotItemOptions
                (pivotTable.Fields[0] as PivotFieldImpl).AddItemOption(0, options);
                (pivotTable.Fields[0] as PivotFieldImpl).AddItemOption(1, options);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ExpandOrCollapse.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

