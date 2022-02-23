using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Remove_Column
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
                IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:C5"]);

                // Removing a column from the table
                worksheet.DeleteColumn(2, 1);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("RemoveTableColumn.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("RemoveTableColumn.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
