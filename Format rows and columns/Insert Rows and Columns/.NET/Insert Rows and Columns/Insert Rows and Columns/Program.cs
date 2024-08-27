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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
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
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/InsertRowsandColumns.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





