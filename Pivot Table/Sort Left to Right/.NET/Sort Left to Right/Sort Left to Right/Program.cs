using System.IO;
using Syncfusion.XlsIO;

namespace Sort_Left_to_Right
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
                IWorksheet sheet = workbook.Worksheets[1];
                IPivotTable pivotTable = sheet.PivotTables[0];

                // Pivot table Left to Right sorting.
                IPivotField columnField = pivotTable.ColumnFields[0];
                columnField.AutoSort(PivotFieldSortType.Descending, 1);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/PivotSort.xlsx"));
                #endregion
            }
        }
    }
}





