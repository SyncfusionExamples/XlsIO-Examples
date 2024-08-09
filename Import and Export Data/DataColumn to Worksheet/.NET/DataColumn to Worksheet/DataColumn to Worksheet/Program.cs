﻿using System;
using Syncfusion.XlsIO;
using System.Data;
using System.IO;

namespace DataColumn_to_Worksheet
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

                #region Import from DataColumn
                //Initialize the DataTable
                DataTable table = SampleDataTable();
                //Import Data Column to the worksheet
                DataColumn column = table.Columns[0];
                worksheet.ImportDataColumn(column, true, 1, 1);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ImportDataColumn.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ImportDataColumn.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
        private static DataTable SampleDataTable()
        {
            //Create a DataTable with four columns
            DataTable table = new DataTable();
            table.Columns.Add("Dosage", typeof(int));
            table.Columns.Add("Drug", typeof(string));
            table.Columns.Add("Patient", typeof(string));
            table.Columns.Add("Date", typeof(DateTime));

            //Add five DataRows
            table.Rows.Add(25, "Indocin", "David", DateTime.Now);
            table.Rows.Add(50, "Enebrel", "Sam", DateTime.Now);
            table.Rows.Add(10, "Hydralazine", "Christoff", DateTime.Now);
            table.Rows.Add(21, "Combivent", "Janet", DateTime.Now);
            table.Rows.Add(100, "Dilantin", "Melanie", DateTime.Now);

            return table;
        }
    }
}
