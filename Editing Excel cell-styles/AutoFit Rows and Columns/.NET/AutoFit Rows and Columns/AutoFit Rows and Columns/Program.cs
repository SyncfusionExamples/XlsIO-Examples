﻿using System.IO;
using Syncfusion.XlsIO;

namespace AutoFit_Rows_and_Columns
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

                #region AutoFit Row
                //Auto-fit rows
                worksheet.Range["A2"].Text = "Fit the content to row";
                worksheet.Range["A2"].WrapText = true;
                worksheet.Range["A2"].AutofitRows();
                #endregion

                #region AutoFit Column
                //Auto-fit columns
                worksheet.Range["B4"].Text = "Fit the content to column";
                worksheet.Range["B4"].AutofitColumns();
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/AutoFit.xlsx"));
                #endregion
            }
        }
    }
}




