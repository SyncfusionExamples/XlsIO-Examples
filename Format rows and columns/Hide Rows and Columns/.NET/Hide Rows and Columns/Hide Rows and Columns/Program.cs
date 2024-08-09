using System.IO;
using Syncfusion.XlsIO;

namespace Hide_Rows_and_Columns
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

                #region Hide Row and Column
                //Hiding the first column and second row
                worksheet.ShowColumn(1, false);
                worksheet.ShowRow(2, false);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("HideRowsandColumns.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("HideRowsandColumns.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
