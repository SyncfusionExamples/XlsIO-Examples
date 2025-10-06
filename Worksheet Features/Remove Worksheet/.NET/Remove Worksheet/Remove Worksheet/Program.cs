using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Worksheet
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

                #region Remove
                //Removing the sheet
                workbook.Worksheets[0].Remove();
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/RemoveWorksheet.xlsx"));
                #endregion
            }
        }
    }
}





