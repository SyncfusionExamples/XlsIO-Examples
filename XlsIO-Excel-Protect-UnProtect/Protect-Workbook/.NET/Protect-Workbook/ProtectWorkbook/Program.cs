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

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputWorkbook.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Protect workbook with password
                workbook.Protect(true, true, "syncfusion");
				
				#region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ProtectedWorkbook.xlsx"));
                #endregion
            }
        }
    }
}