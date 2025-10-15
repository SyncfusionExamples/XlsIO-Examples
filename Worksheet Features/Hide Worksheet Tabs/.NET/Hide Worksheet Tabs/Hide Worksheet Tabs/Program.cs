using System.IO;
using Syncfusion.XlsIO;

namespace Hide_Worksheet_Tabs
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
                sheet.Range["A1:M20"].Text = "Tabs";

                #region Hide Worksheet Tabs
                //Hide the tab
                workbook.DisplayWorkbookTabs = false;
                //set the display tab
                workbook.DisplayedTab = 2;
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/HideWorksheetTabs.xlsx"));
                #endregion
            }
        }
    }
}




