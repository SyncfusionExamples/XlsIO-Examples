using System.IO;
using Syncfusion.XlsIO;

namespace Move_Rows_and_Columns
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

                #region Move Rows
                //Shifts cells toward Up after deletion
                worksheet.Range["A4:A8"].Clear(ExcelMoveDirection.MoveUp);
                #endregion

                #region Move Columns
                //Shifts cells towards Left after deletion
                worksheet.Range["B1:E1"].Clear(ExcelMoveDirection.MoveLeft);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("MoveRowsandColumns.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("MoveRowsandColumns.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
