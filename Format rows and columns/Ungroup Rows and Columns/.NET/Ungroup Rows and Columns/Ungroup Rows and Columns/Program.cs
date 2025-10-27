using System.IO;
using Syncfusion.XlsIO;

namespace Ungroup_Rows_and_Columns
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate - ToUngroup.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Un-Group Rows
                //Ungroup Rows
                worksheet.Range["A3:A7"].Ungroup(ExcelGroupBy.ByRows);
                #endregion

                #region Un-Group Columns
                //Ungroup Columns
                worksheet.Range["C1:D1"].Ungroup(ExcelGroupBy.ByColumns);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/UngroupRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





