using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Read_Table
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

                //Accessing first table in the sheet
                IListObject table = worksheet.ListObjects[0];

                //Modifying table name
                table.Name = "SalesTable";

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ReadTable.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ReadTable.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
