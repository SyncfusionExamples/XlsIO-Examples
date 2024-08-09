﻿using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Decrypt_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/EncryptedWorkbook.xlsx", FileMode.Open, FileAccess.ReadWrite);
				
                //Open encrypted Excel document with password
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelParseOptions.Default, false, "syncfusion");
                IWorksheet worksheet = workbook.Worksheets[0];

                //Decrypt workbook
                workbook.PasswordToOpen = string.Empty;
				
				#region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("DecryptedWorkbook.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
				inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("DecryptedWorkbook.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
