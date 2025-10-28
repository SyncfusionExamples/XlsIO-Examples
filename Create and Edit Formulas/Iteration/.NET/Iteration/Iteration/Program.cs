﻿using System.IO;
using Syncfusion.XlsIO;

namespace Iteration
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

                //Setting iteration
                workbook.CalculationOptions.IsIterationEnabled = true;

                //Number of times to recalculate
                workbook.CalculationOptions.MaximumIteration = 99;

                //Number of acceptable changes
                workbook.CalculationOptions.MaximumChange = 40;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Iteration.xlsx"));
                #endregion
            }
        }
    }
}




