using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Decrypt_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
				
                //Open encrypted Excel document with password
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/EncryptedWorkbook.xlsx"), ExcelParseOptions.Default, false, "syncfusion");
                IWorksheet worksheet = workbook.Worksheets[0];

                //Decrypt workbook
                workbook.PasswordToOpen = string.Empty;
				
				#region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/DecryptedWorkbook.xlsx"));
                #endregion
            }
        }
    }
}





