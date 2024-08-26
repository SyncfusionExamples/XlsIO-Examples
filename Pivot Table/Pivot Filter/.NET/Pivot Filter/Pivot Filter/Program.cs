using System.IO;
using Syncfusion.XlsIO;

namespace Pivot_Filter
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

                IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:H50"]);

                IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);
                pivotTable.Fields[4].Axis = PivotAxisTypes.Page;
                pivotTable.Fields[2].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[6].Axis = PivotAxisTypes.Row;
                pivotTable.Fields[3].Axis = PivotAxisTypes.Column;

                IPivotField dataField = pivotSheet.PivotTables[0].Fields[5];
                pivotTable.DataFields.Add(dataField, "Sum of Units", PivotSubtotalTypes.Sum);

                //Apply page filter
                pivotTable.Fields[4].Axis = PivotAxisTypes.Page;

                IPivotField pageField = pivotTable.Fields[4];

                pageField.Items[1].Visible = false;
                pageField.Items[2].Visible = false;

                //Apply label filter
                IPivotField rowField = pivotTable.Fields[2];
                rowField.PivotFilters.Add(PivotFilterType.CaptionEqual, null, "East", null);

                //Apply item filter
                IPivotField colField = pivotTable.Fields[3];
                colField.Items[0].Visible = false;
                colField.Items[1].Visible = false;

                //Apply value filter
                IPivotField field = pivotTable.Fields[2];
                field.PivotFilters.Add(PivotFilterType.ValueLessThan, field, "1341", null);

                pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium2;
                pivotSheet.Activate();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/PivotFilter.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

