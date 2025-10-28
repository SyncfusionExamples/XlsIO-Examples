﻿using System.IO;
using Syncfusion.XlsIO;

namespace Incremental_Formula
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Enables the incremental formula to updates the reference in cell
                application.EnableIncrementalFormula = true;

                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Formula are automatically increments by one for the range of cells
                sheet["A1:A5"].Formula = "=B1+C1";

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/IncrementalFormula.xlsx"));
                #endregion
            }
        }
    }
}




