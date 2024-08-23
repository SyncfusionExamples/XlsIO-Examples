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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputData.xlsx"), FileMode.Open, FileAccess.ReadWrite);
 
                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Unlock cell
                worksheet["A1"].CellStyle.Locked = false;
				
				#region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("LockedCells.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
				inputStream.Dispose();
            }
        }
    }
}

