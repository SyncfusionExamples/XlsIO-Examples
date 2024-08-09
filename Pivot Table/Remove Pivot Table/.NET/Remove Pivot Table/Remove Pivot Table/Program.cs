using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Pivot_Table
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
                IWorksheet pivotSheet = workbook.Worksheets[1];

                //Remove the pivot table
                pivotSheet.PivotTables.Remove("PivotTable1");

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
