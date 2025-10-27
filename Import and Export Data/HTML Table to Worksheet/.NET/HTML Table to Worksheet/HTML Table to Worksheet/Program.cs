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
                workbook.SaveAs(Path.GetFullPath("Output/HTMLTabletoWorksheet.xlsx"));
                #endregion

                //Dispose streams
                inputStream.Dispose();
            }
        }
    }
}





