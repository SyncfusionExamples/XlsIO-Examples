using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace Column_Width_For_Pivot_Table_Range
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
                IWorksheet worksheet1 = workbook.Worksheets[1];

                //Create pivot cache with the given data range
                IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:H5"]);

                //Create pivot table with the cache at the specified range
                IPivotTable pivotTable = worksheet1.PivotTables.Add("PivotTable1", worksheet1["A1"], cache);

                PivotTableImpl pivotTableImpl = pivotTable as PivotTableImpl;

                //Add Pivot table fields 
                pivotTable.Fields[0].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[1].Axis = PivotAxisTypes.Row;
                pivotTable.DataFields.Add(pivotTable.Fields["Total"], "Sum", PivotSubtotalTypes.Sum);

                //Set column width
                worksheet1.Range["A10"].ColumnWidth = 50;

                //Disable pivot table autoformat    
                (pivotTable.Options as PivotTableOptions).IsAutoFormat = false;

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
