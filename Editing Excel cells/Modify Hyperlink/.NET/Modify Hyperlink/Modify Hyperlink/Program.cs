﻿using System.IO;
using Syncfusion.XlsIO;

namespace Modify_Hyperlink
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

                #region Modify Hyperlink
                //Modifying hyperlink’s text to display
                IHyperLink hyperlink = worksheet.Range["C5"].Hyperlinks[0];
                hyperlink.TextToDisplay = "Syncfusion";
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ModifyHyperlink.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ModifyHyperlink.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
