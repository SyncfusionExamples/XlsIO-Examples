﻿using System.IO;
using Syncfusion.XlsIO;

namespace Paste_As_Link
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Paste As Link
                //Copy range as link from Range “A1” to “A5”
                IRange source = worksheet.Range["A1"];
                IRange destination = worksheet.Range["A5"];
                source.CopyTo(destination, true);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("PasteAsLink.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PasteAsLink.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
