using System.IO;
using Syncfusion.XlsIO;

namespace Format_Pivot_Cell
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
                //Get the cell format for pivot range.
                IPivotCellFormat cellFormat = pivotTable.GetCellFormat("A4:J5");
                cellFormat.BackColor = ExcelKnownColors.Green;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/PivotCellFormat.xlsx"));
                #endregion
            }
        }
    }
}





