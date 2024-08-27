using Syncfusion.XlsIO;
using System;
using System.IO;

namespace ProtectWorkbook
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
				application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputWorkbook.xlsx"), FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                workbook.Unprotect();
                //Protect workbook with password
                workbook.Protect(true, true, "syncfusion");
				
				#region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ProtectedWorkbook.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
				inputStream.Dispose();
            }
        }
    }
}





