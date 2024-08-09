﻿using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Conditional_Format
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Removing conditional format for a specified range 
                worksheet.Range["E5"].ConditionalFormats.Remove();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("RemoveConditionalFormat.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("RemoveConditionalFormat.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
