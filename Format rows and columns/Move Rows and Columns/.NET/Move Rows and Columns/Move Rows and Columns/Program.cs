using System.IO;
using Syncfusion.XlsIO;

namespace Move_Rows_and_Columns
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
                
                #region Move Rows
                //Shifts cells toward Up after deletion
                worksheet.Range["A4:A8"].Clear(ExcelMoveDirection.MoveUp);
                #endregion

                #region Move Columns
                //Shifts cells towards Left after deletion
                worksheet.Range["B1:E1"].Clear(ExcelMoveDirection.MoveLeft);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/MoveRowsandColumns.xlsx"));
                #endregion
            }
        }
    }
}





