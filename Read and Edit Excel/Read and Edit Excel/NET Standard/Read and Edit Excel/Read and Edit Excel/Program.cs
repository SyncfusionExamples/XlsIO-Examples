using System.IO;
using Syncfusion.XlsIO;

namespace Read_and_Edit_Excel
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            //Creates a new instance for ExcelEngine, and will Close on exiting using block
            using (var excelEngine = new ExcelEngine())
            {
                IWorkbook workbook;     // N.B. IWorkbook is not IDisposable so can't wrap in using, hence explicit Close below

                #region Open
                // open an existing workbook through Open method of IWorkbooks, needs explicit ExcelVersion if writing (see below)
                // N.B. input .xlsx file was copied to output folder at build time
                using (var inputStream = new FileStream("InputTemplate.xlsx", FileMode.Open, FileAccess.Read))
                {
                    workbook = excelEngine.Excel.Workbooks.Open(inputStream, ExcelVersion.Xlsx);
                }
                #endregion

                #region Edit
                //Set a value in Excel cell
                workbook.Worksheets[0].Range["A2"].Value = "Hello World";
                #endregion

                #region Save
                //Saving the workbook
                using (var outputStream = new FileStream("ReadandEditExcel.xlsx", FileMode.Create, FileAccess.Write))
                {
                    workbook.SaveAs(outputStream);
                }
                #endregion

                #region Close
                //Close the instance of IWorkbook
                workbook.Close();
                #endregion

            }
            var process = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo("ReadandEditExcel.xlsx")
                {
                    UseShellExecute = true
                }
            };
            _ = process.Start();
        }
    }
}
