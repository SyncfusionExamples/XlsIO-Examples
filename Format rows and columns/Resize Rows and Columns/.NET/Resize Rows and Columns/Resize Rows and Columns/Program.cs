using System.IO;
using Syncfusion.XlsIO;

namespace Resize_Rows_and_Columns
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
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Resize rows
                //Modifying the row height of one row
                worksheet.SetRowHeight(2, 100);
                
                //Modifying the row height of multiple rows
                worksheet.Range["A5:A10"].RowHeight = 40;
                #endregion

                #region Resize columns
                //Modifying the column width of one column
                worksheet.SetColumnWidth(2, 50);

                //Modifying the column width of multiple columns
                worksheet.Range["D1:G1"].ColumnWidth = 5;
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ResizeRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





