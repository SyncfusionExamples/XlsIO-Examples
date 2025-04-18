﻿using System.IO;
using Syncfusion.XlsIO;

namespace Format_Pivot_Table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[1];
                IPivotTable pivotTable = worksheet.PivotTables[0];

                //Set BuiltInStyle
                pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleDark12;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/FormatPivotTable.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





