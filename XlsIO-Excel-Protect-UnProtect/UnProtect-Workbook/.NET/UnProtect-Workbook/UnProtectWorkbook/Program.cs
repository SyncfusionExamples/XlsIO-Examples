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

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputWorkbook.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //UnProtect workbook with password
                workbook.Unprotect("syncfusion");
				
				#region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/UnProtectedWorkbook.xlsx"));
                #endregion
            }
        }
    }
}





