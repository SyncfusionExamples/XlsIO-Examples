
using Syncfusion.XlsIO;

namespace PivotTable
{
    class Program
    {
        public static void Main()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/SalesReport.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                IWorksheet pivotSheet = workbook.Worksheets.Create("PivotSheet");

                //Create Pivot cache with the given data range
                IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:H50"]);

                //Create "PivotTable1" with the cache at the specified range
                IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);

                //Add Pivot table row fields                
                pivotTable.Fields[3].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[4].Axis = PivotAxisTypes.Row;

                //Add Pivot table column fields
                pivotTable.Fields[2].Axis = PivotAxisTypes.Column;

                //Add data fields
                IPivotField field = pivotTable.Fields[5];
                pivotTable.DataFields.Add(field, "Units", PivotSubtotalTypes.Sum);

                field = pivotTable.Fields[6];
                pivotTable.DataFields.Add(field, "Unit Cost", PivotSubtotalTypes.Sum);

                //Pivot table style
                pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium14;    
                
                pivotSheet.Activate();

                string fileName = "PivotTable.xlsx";
                //Saving the workbook as stream
                FileStream stream = new FileStream(fileName, FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }
    }
}

