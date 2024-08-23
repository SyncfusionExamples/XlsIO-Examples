using System.IO;
using Syncfusion.XlsIO;

namespace HTML_Table_to_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Imports HTML table into the worksheet from first row and first column
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.html"), FileMode.Open, FileAccess.ReadWrite);
                worksheet.ImportHtmlTable(inputStream, 1, 1);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("HTMLTabletoWorksheet.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

