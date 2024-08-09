﻿using System.IO;
using Syncfusion.XlsIO;

namespace Wrap_Text
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["A2"].Text = "First Sentence is wrapped";
                worksheet.Range["B2"].Text = "Second Sentence is wrapped";
                worksheet.Range["C2"].Text = "Third Sentence is wrapped";

                #region Wrap Text
                //Applying Wrap-text
                worksheet.Range["A2:C2"].WrapText = true;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("WrapText.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("WrapText.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
