﻿using System.IO;
using Syncfusion.XlsIO;

namespace Dynamic_Refresh
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet pivotSheet = workbook.Worksheets[0];

                //Change the range values that the Pivot Tables range refers to
                workbook.Names["PivotRange"].RefersToRange = pivotSheet.Range["A1:H25"];

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("PivotTable.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PivotTable.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
