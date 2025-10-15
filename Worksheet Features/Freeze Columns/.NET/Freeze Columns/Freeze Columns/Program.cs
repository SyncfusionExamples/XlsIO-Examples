﻿using System.IO;
using Syncfusion.XlsIO;

namespace Freeze_Columns
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Applying freeze columns to the sheet by specifying a cell
                worksheet.Range["C1"].FreezePanes();

                //Set first visible column in the right pane
                worksheet.FirstVisibleColumn = 4;

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}





