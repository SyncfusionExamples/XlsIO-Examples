using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Override_Excel_Document
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Open an existing Excel file as stream
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Sample.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Modify the data
                worksheet.Range["A1"].Text = "Hello World";

                //Dispose input stream
                inputStream.Dispose();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Sample.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose output stream
                outputStream.Dispose();
            }
        }
    }
}
