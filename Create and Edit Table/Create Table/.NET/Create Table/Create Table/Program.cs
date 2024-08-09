using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Create_Table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];
                
                //Create for the given data
                IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:C5"]);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("CreateTable.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("CreateTable.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
