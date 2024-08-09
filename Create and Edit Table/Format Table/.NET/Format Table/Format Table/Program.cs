using Syncfusion.XlsIO;
using System;
using System.IO;

namespace Format_Table
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

                //Formatting table with a built-in style
                table.BuiltInTableStyle = TableBuiltInStyles.TableStyleMedium9;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("FormatTable.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("FormatTable.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
