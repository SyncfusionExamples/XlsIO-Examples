using System.IO;
using Syncfusion.XlsIO;

namespace Group_Rows_and_Columns
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate - ToGroup.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Group Rows
                //Group Rows
                worksheet.Range["A3:A7"].Group(ExcelGroupBy.ByRows, true);
                worksheet.Range["A11:A16"].Group(ExcelGroupBy.ByRows);
                #endregion

                #region Group Columns
                //Group Columns
                worksheet.Range["C1:D1"].Group(ExcelGroupBy.ByColumns, false);
                worksheet.Range["F1:G1"].Group(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/GroupRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





