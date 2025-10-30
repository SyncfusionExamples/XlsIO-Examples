using System.IO;
using Syncfusion.XlsIO;

namespace Sort_Top_to_Bottom
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

                // Pivot Top to Bottom sorting.
                IPivotField rowField = pivotTable.RowFields[0];
                rowField.AutoSort(PivotFieldSortType.Ascending, 1);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/PivotSort.xlsx"));
                #endregion
            }
        }
    }
}





