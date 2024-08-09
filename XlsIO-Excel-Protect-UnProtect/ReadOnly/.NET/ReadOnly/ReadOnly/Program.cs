using Syncfusion.XlsIO;
using System;
using System.IO;

namespace ReadOnly
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
				application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputData.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Set Read only
                workbook.ReadOnlyRecommended = true;
				
				#region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ReadOnly.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
				inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ReadOnly.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
