using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;
using System.IO;


namespace Refresh_Pivot_Chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2013;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/PivotChart.xlsx"));

                IWorksheet dataSheet = workbook.Worksheets[0];
                IWorksheet pivotSheet = workbook.Worksheets[1];

                // Update pivot cache source range to refresh the PivotChart
                (pivotSheet.PivotTables[0] as PivotTableImpl).Cache.SourceRange = dataSheet["A1:H50"];

                workbook.SaveAs(Path.GetFullPath(@"Output/PivotChart_Refreshed.xlsx"));
            }
        }
    }
}





