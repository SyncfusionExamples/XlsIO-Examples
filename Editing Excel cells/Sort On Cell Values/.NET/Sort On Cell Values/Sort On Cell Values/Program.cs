﻿using System.IO;
using Syncfusion.XlsIO;

namespace Sort_On_Cell_Values
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Sort On Cell Values
                //Creates the data sorter
                IDataSort sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.Range["A1:A11"];

                //Adds the sort field with the column index, sort based on and order by attribute
                sorter.SortFields.Add(0, SortOn.Values, OrderBy.Ascending);

                //Sort based on the sort Field attribute
                sorter.Sort();

                //Creates the data sorter
                sorter = workbook.CreateDataSorter();

                //Range to sort
                sorter.SortRange = worksheet.Range["B1:B11"];

                //Adds the sort field with the column index, sort based on and order by attribute
                sorter.SortFields.Add(1, SortOn.Values, OrderBy.Descending);

                //Sort based on the sort Field attribute
                sorter.Sort();
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("SortOnValues.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("SortOnValues.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
