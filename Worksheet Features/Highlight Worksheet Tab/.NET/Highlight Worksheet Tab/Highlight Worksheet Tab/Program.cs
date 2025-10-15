using System.IO;
using Syncfusion.XlsIO;

namespace Highlight_Worksheet_Tab
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

                #region Highlight Worksheet Tab
                //Highlighting sheet tab
                sheet.TabColor = ExcelKnownColors.Green;
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/HighlightSheetTab.xlsx"));
                #endregion
            }
        }
    }
}




