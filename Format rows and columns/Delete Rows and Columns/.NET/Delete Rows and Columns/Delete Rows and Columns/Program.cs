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
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
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
                FileStream outputStream = new FileStream("DeleteRowsandColumns.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("DeleteRowsandColumns.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
