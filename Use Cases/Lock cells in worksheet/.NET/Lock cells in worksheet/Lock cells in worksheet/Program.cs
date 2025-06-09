using Syncfusion.XlsIO;

namespace Lock_Cells_In_Worksheet
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputData.xlsx"), FileMode.Open, FileAccess.ReadWrite);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Unlock cells in the worksheet
                worksheet.UsedRange.CellStyle.Locked = false; 

                //Lock specific cells (A1:A5)
                worksheet.Range["A1:A5"].CellStyle.Locked = true;

                //Protect worksheet to apply lock
                worksheet.Protect("password", ExcelSheetProtection.All);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("LockedCells.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}