using System.IO;
using Syncfusion.XlsIO;

namespace Create_Access_Remove_Worksheet
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

                #region  Access
                //Accessing via index
                IWorksheet sheet = workbook.Worksheets[0];

                //Accessing via sheet name
                IWorksheet NamedSheet = workbook.Worksheets["Sample"];
                #endregion

                #region Remove
                //Removing the sheet
                workbook.Worksheets[0].Remove();
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("CreateAccessRemove.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CreateAccessRemove.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
