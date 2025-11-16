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

                //Open an existing Excel file 
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Sample.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Modify the data
                worksheet.Range["A1"].Text = "Hello World";

                //Dispose input stream
                inputStream.Dispose();

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Sample.xlsx"));
                #endregion
            }
        }
    }
}
