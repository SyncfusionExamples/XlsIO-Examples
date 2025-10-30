using System;
using System.IO;
using Syncfusion.XlsIO;

namespace Create_Table
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
                
                //Create for the given data
                IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:C5"]);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CreateTable.xlsx"));
                #endregion
            }
        }
    }
}