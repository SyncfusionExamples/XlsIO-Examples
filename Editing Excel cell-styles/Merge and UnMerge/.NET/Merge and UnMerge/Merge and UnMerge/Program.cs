﻿using System.IO;
using Syncfusion.XlsIO;

namespace Merge_and_UnMerge
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

                #region Merge
                //Merging cells
                worksheet.Range["A5:E10"].Merge();
                worksheet.Range["A15:E20"].Merge();
                #endregion

                #region UnMerge
                //Un-Merging merged cells
                worksheet.Range["A5:E10"].UnMerge();
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("MergeandUnMerge.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("MergeandUnMerge.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
