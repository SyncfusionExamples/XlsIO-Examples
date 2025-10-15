using System.IO;
using Syncfusion.XlsIO;

namespace Hide_Gridlines
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
                sheet.Range["A1:M20"].Text = "Gridlines";

                #region Hide Gridlines
                //Hide grid line
                sheet.IsGridLinesVisible = false;
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/HideGridlines.xlsx"));
                #endregion
            }
        }
    }
}




