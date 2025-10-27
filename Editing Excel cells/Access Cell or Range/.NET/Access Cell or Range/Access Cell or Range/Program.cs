using System.IO;
using Syncfusion.XlsIO;

namespace Access_Cell_or_Range
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

                #region Access Access Cell or Range
                //Access a range by specifying cell address
                sheet.Range["A7"].Text = "Accessing a Range by specify cell address ";

                //Access a range by specifying cell row and column index
                sheet.Range[9, 1].Text = "Accessing a Range by specify cell row and column index ";

                //Access a Range by specifying using defined name
                IName name = workbook.Names.Add("Name");
                name.RefersToRange = sheet.Range["A11"];
                sheet.Range["Name"].Text = "Accessing a Range by specifying using defined name";

                //Accessing a Range of cells by specifying cells address
                sheet.Range["A13:C13"].Text = "Accessing a Range of Cells (Method 1)";

                //Accessing a Range of cells specifying cell row and column index
                sheet.Range[15, 1, 15, 3].Text = "Accessing a Range of Cells (Method 2)";
                #endregion

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/AccessCellorRange.xlsx"));
                #endregion
            }
        }
    }
}




