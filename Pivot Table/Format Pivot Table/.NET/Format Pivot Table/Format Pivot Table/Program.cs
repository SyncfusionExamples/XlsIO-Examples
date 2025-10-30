using System.IO;
using Syncfusion.XlsIO;

namespace Format_Pivot_Table
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
                IWorksheet worksheet = workbook.Worksheets[1];
                IPivotTable pivotTable = worksheet.PivotTables[0];

                //Set BuiltInStyle
                pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleDark12;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/FormatPivotTable.xlsx"));
                #endregion
            }
        }
    }
}





