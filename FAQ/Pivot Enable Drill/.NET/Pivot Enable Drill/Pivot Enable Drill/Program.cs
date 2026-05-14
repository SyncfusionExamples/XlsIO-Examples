using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

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

            // How to enable drilldown in pivot table
            (pivotTable as PivotTableImpl).EnableDrilldown = true;

            #region Save
            //Saving the workbook
            workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            #endregion
        }
    }
}
