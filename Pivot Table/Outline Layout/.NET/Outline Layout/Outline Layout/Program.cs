using Syncfusion.XlsIO;
using System.IO;

namespace OutLine_Layout
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[1];
                IPivotTable pivotTable = worksheet.PivotTables[0];

                //Set PivotTableRowLayout
                pivotTable.Options.RowLayout = PivotTableRowLayout.Outline;

                //Set BuiltInStyle
                pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium9;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}




