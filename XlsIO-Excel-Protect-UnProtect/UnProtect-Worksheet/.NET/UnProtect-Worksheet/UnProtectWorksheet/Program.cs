﻿using Syncfusion.XlsIO;
using System;
using System.IO;

namespace UnProtectWorksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
				application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/ProtectedWorksheet.xlsx", FileMode.Open, FileAccess.ReadWrite);

                //Open Excel
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //UnProtect worksheet with password
                worksheet.Unprotect("syncfusion");
				
				#region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("UnProtectedSheet.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
				inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("UnProtectedSheet.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
