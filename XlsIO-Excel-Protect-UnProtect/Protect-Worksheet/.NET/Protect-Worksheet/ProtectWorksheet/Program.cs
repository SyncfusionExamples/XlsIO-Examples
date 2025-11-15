using Syncfusion.XlsIO;
using System;
using System.IO;

namespace ProtectWorksheet
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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputData.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Protect worksheet with multiple options
                worksheet.Protect("Protect", ExcelSheetProtection.FormattingCells | ExcelSheetProtection.LockedCells | ExcelSheetProtection.UnLockedCells);
                				
				#region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ProtectedSheet.xlsx"));
                #endregion
            }
        }
    }
}





