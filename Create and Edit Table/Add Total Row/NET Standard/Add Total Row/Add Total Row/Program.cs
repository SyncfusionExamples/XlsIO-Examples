using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Add_Total_Row
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Creating a table
                IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:C8"]);

                //Adding total row
                table.ShowTotals = true;
                table.Columns[0].TotalsRowLabel = "Total";
                table.Columns[1].TotalsCalculation = ExcelTotalsCalculation.Sum;
                table.Columns[2].TotalsCalculation = ExcelTotalsCalculation.Sum;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("AddTotalRow.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("AddTotalRow.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
