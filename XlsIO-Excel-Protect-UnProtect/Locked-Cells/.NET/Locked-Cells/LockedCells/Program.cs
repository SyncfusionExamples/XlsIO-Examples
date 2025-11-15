using System;
using System.IO;
using Syncfusion.XlsIO;

namespace LockedCells
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

                //Unlock cell
                worksheet["A1"].CellStyle.Locked = false;
				
				#region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/LockedCells.xlsx"));
                #endregion
            }
        }
    }
}





