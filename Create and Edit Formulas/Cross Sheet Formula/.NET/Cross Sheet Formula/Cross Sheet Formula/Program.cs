﻿using System.IO;
using Syncfusion.XlsIO;

namespace Cross_Sheet_Formula
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(2);
                IWorksheet sheet1 = workbook.Worksheets[0];
                IWorksheet sheet2 = workbook.Worksheets[1];

                sheet1.SetValue(2, 1, "20");
                sheet2.SetValue(2, 2, "10");

                #region Cross Sheet Formula
                //Setting formula for the range with cross-sheet reference
                sheet1.Range["C2"].Formula = "=SUM(Sheet2!B2,Sheet1!A2)";
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Formula.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




