using System.IO;
using Syncfusion.XlsIO;

namespace Autofit_Rows_and_Columns
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

                #region Autofit Rows
                //Autofit applied to a single row
                worksheet.AutofitRow(3);

                //Autofit applied to multiple rows
                worksheet.Range["6:10"].AutofitRows();
                #endregion

                #region Autofit Columns
                //Autofit applied to a single column
                worksheet.AutofitColumn(2);

                //Autofit applied to multiple columns
                worksheet.Range["E:G"].AutofitColumns();
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/AutofitRowsandColumns.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("AutofitRowsandColumns.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
