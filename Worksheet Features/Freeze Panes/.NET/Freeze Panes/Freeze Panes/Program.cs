using System.IO;
using Syncfusion.XlsIO;

namespace Freeze_Panes
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
                IWorksheet sheet = workbook.Worksheets[0];

                #region Freeze Panes
                //Applying Freeze Pane to the sheet by specifying a cell
                sheet.Range["B2"].FreezePanes();
                #endregion

                #region First Visible Row
                //Set first visible row in the bottom pane
                sheet.FirstVisibleRow = 2;
                #endregion

                #region First Visible Column
                //Set first visible column in the right pane
                sheet.FirstVisibleColumn = 2;
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/FreezePanes.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




