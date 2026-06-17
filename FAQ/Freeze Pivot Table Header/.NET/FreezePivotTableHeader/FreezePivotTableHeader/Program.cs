using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace FreezePivotTableHeader
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

                //Freeze row and column
                IWorksheet freezeSheet = workbook.Worksheets[1];
                IRange range = freezeSheet.PivotTables[0].Location;
                freezeSheet[range.Row + 1, range.Column + 1].FreezePanes();

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}





