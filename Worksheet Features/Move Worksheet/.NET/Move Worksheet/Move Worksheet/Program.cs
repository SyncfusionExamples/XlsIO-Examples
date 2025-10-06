using System.IO;
using Syncfusion.XlsIO;

namespace Move_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Create(3);
                IWorksheet sheet = workbook.Worksheets[0];

                #region Move Worksheet
                //Move the Sheet
                sheet.Move(1);
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/MoveWorksheet.xlsx"));
                #endregion
            }
        }
    }
}




