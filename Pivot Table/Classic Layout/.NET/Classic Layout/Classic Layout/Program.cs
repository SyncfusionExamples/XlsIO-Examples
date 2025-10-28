using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace Classic_Layout
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

                //Set classic layout
                (pivotTable.Options as PivotTableOptions).ShowGridDropZone = true;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ClassicLayout.xlsx"));
                #endregion
            }
        }
    }
}





