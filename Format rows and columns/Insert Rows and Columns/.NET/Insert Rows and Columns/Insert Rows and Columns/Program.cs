using System.IO;
using Syncfusion.XlsIO;

namespace Insert_Rows_and_Columns
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

                #region Insert Rows
                //Insert a row
                worksheet.InsertRow(3, 1, ExcelInsertOptions.FormatAsBefore);

                //Insert multiple rows
                worksheet.InsertRow(10, 3, ExcelInsertOptions.FormatAsAfter);
                #endregion

                #region Insert Columns
                //Insert a column
                worksheet.InsertColumn(2, 1, ExcelInsertOptions.FormatAsAfter);

                //Insert multiple columns
                worksheet.InsertColumn(9, 2, ExcelInsertOptions.FormatAsBefore);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/InsertRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





