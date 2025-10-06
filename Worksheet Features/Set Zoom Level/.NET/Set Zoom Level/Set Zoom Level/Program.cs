using System.IO;
using Syncfusion.XlsIO;

namespace Set_Zoom_Level
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
                sheet.Range["A1:M20"].Text = "Zoom level";

                #region Set Zoom Level
                //set zoom percentage
                sheet.Zoom = 70;
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/SetZoomLevel.xlsx"));
                #endregion
            }
        }
    }
}




