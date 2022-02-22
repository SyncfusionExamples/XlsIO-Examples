using System.IO;
using Syncfusion.XlsIO;

namespace Pivot_Layout
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //The first worksheet object in the worksheets collection is accessed.
                IWorksheet sheet = workbook.Worksheets[0];
                //Access the sheet to draw pivot table.
                IWorksheet pivotSheet = workbook.Worksheets[1];
                pivotSheet.Activate();

                //Select the data to add in cache
                IPivotCache cache = workbook.PivotCaches.Add(sheet["A1:G20"]);
                //Insert the pivot table. 
                IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);
                pivotTable.Fields[0].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[1].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[2].Axis = PivotAxisTypes.Row;
                IPivotField field1 = pivotSheet.PivotTables[0].Fields[5];
                pivotTable.DataFields.Add(field1, "Sum of Land Area", PivotSubtotalTypes.Sum);
                IPivotField field2 = pivotSheet.PivotTables[0].Fields[6];
                pivotTable.DataFields.Add(field2, "Sum of Water Area", PivotSubtotalTypes.Sum);

                //Select RowLayout. Available options are Outline, Tabular and Compact
                pivotTable.Options.RowLayout = PivotTableRowLayout.Outline;
                pivotTable.Location = pivotSheet.Range[1, 1, 51, 5];

                //Apply built in style.
                pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium9;
                pivotSheet.Range[1, 1, 1, 14].ColumnWidth = 11;
                pivotSheet.SetColumnWidth(1, 15.29);
                pivotSheet.SetColumnWidth(2, 15.29);

                //Layout the pivot table.
                pivotTable.Layout();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("PivotLayout.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PivotLayout.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
