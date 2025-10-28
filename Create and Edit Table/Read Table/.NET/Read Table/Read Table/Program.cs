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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Accessing first table in the sheet
                IListObject table = worksheet.ListObjects[0];

                //Modifying table name
                table.DisplayName = "SalesTable";

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ReadTable.xlsx"));
                #endregion
            }
        }
    }
}





