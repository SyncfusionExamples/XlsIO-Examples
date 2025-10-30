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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
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
                workbook.SaveAs(Path.GetFullPath("Output/AddTotalRow.xlsx"));
                #endregion
            }
        }
    }
}





