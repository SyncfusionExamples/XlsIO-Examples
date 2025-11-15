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

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputExcel.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Encrypt workbook with password
                workbook.PasswordToOpen = "syncfusion";                
				
				#region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/EncryptedWorkbook.xlsx"));
                #endregion
            }
        }
    }
}





