using System.IO;
using Syncfusion.XlsIO;

namespace Hide_Rows_and_Columns
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

                #region Hide Row and Column
                //Hiding the first column and second row
                worksheet.ShowColumn(1, false);
                worksheet.ShowRow(2, false);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/HideRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





