using System.IO;
using Syncfusion.XlsIO;

namespace Create_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                #region Create
                //The new workbook is created with 5 worksheets
                IWorkbook workbook = application.Workbooks.Create(5);
                //Creating a new sheet
                IWorksheet worksheet = workbook.Worksheets.Create();
                //Creating a new sheet with name “Sample”
                IWorksheet namedSheet = workbook.Worksheets.Create("Sample");
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("CreateWorksheet.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CreateWorksheet.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
