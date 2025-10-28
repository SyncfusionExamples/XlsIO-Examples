using System.IO;
using Syncfusion.XlsIO;

namespace Formula_Array
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

                #region Formula Array
                //Assign array formula
                sheet.Range["A1:D1"].FormulaArray = "{1,2,3,4}";

                //Adding a named range for the range A1 to D1
                sheet.Names.Add("ArrayRange", sheet.Range["A1:D1"]);

                //Assign formula array with named range
                sheet.Range["A2:D2"].FormulaArray = "ArrayRange+100";
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Formula.xlsx"));
                #endregion
            }
        }
    }
}




