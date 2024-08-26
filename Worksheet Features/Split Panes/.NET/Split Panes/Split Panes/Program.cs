using System.IO;
using Syncfusion.XlsIO;

namespace Split_Panes
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

                #region Split Panes
                //split panes
                sheet.FirstVisibleColumn = 2;
                sheet.FirstVisibleRow = 5;
                sheet.VerticalSplit = 5000;
                sheet.HorizontalSplit = 5000;
                #endregion

                sheet.ActivePane = 1;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/SplitPanes.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




