﻿using System.IO;
using Syncfusion.XlsIO;

namespace Array_to_Worksheet
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

                //Initialize the Object Array
                object[] array = new object[4] { "Total Income", "Actual Expense", "Expected Expenses", "Profit" };
                //Import the Object Array to Sheet
                worksheet.ImportArray(array, 1, 1, false);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ArrayToWorksheet.xlsx"));
                #endregion
            }
        }
    }
}




