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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Creating a table
                IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:C5"]);

                // Removing a column from the table
                worksheet.DeleteColumn(2, 1);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveTableColumn.xlsx"));
                #endregion
            }
        }
    }
}





