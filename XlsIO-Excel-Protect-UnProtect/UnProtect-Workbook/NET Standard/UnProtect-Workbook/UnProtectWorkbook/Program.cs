using Syncfusion.XlsIO;
using System;
using System.IO;

namespace UnProtectWorkbook
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
				application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputWorkbook.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //UnProtect workbook with password
                workbook.Unprotect("syncfusion");
				
				#region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("UnProtectedWorkbook.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
				inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("UnProtectedWorkbook.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
