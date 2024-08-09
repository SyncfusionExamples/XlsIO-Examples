using Syncfusion.Office;
using Syncfusion.XlsIO;
using System;
using System.IO;


namespace Skip_Macro_and_Save
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xls", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Skip Macros while saving
                application.SkipOnSave = SkipExtRecords.Macros;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("SkipMacroAndSave.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream, ExcelSaveType.SaveAsXLS);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("SkipMacroAndSave.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
