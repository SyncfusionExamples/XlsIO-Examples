using System.IO;
using Syncfusion.XlsIO;

namespace Ignore_Error
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
                IWorksheet sheet = workbook.Worksheets[0];

                //Sets warning if number is entered as text.
                sheet.Range["A2:D2"].IgnoreErrorOptions = ExcelIgnoreError.NumberAsText;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/FormulaAuditing.xlsx"));
                #endregion
            }
        }
    }
}





