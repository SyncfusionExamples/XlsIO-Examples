using System.IO;
using Syncfusion.XlsIO;

namespace Activate_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(2);
                IWorksheet sheet = workbook.Worksheets[1];

                sheet.Range["A1:M20"].Text = "Activate";

                #region Activate Worksheet
                //Activate the sheet
                sheet.Activate();
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ActivateWorksheet.xlsx"));
                #endregion
            }
        }
    }
}




