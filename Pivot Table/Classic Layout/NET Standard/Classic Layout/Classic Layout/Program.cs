using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace Classic_Layout
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

                //Set classic layout
                (pivotTable.Options as PivotTableOptions).ShowGridDropZone = true;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ClassicLayout.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ClassicLayout.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
