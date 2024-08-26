using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Encrypt_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
				application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputExcel.xlsx"), FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Encrypt workbook with password
                workbook.PasswordToOpen = "syncfusion";                
				
				#region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/EncryptedWorkbook.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
				inputStream.Dispose();
            }
        }
    }
}





