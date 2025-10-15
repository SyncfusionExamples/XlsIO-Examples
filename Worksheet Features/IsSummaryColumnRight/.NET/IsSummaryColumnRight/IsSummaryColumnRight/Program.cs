﻿using System.IO;
using Syncfusion.XlsIO;

namespace IsSummaryColumnRight
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
                IWorksheet sheet = workbook.Worksheets[0];

                for (int i = 1; i <= 50; i++)
                {
                    for (int j = 1; j <= 50; j++)
                    {
                        sheet.Range[i, j].Text = sheet.Range[i, j].AddressLocal;
                    }
                }

                #region PageSetup Settings
                //True to summary columns will appear right of the detail in outlines
                sheet.PageSetup.IsSummaryColumnRight = true;
                sheet.PageSetup.Orientation = ExcelPageOrientation.Portrait;
                sheet.PageSetup.FitToPagesTall = 0;
                sheet.PageSetup.IsFitToPage = true;

                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/SummaryColumnRight.xlsx"));
                #endregion
            }
        }
    }
}




