using System.IO;
using Syncfusion.XlsIO;

namespace Format_Pivot_Cell
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
                IWorksheet worksheet = workbook.Worksheets[1];

                IPivotTable pivotTable = worksheet.PivotTables[0];
                //Get the cell format for pivot range.
                IPivotCellFormat cellFormat = pivotTable.GetCellFormat("A4:J5");
                cellFormat.BackColor = ExcelKnownColors.Green;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("PivotCellFormat.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PivotCellFormat.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
