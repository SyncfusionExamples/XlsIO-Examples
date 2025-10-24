using System.IO;
using Syncfusion.XlsIO;

namespace Delete_Rows_and_Columns
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

                #region Delete Rows
                //Delete a row
                worksheet.DeleteRow(3);

                //Delete multiple rows
                worksheet.DeleteRow(10, 3);
                #endregion

                #region Delete Columns
                //Delete a column
                worksheet.DeleteColumn(2);

                //Delete multiple columns
                worksheet.DeleteColumn(3, 2);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/DeleteRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





