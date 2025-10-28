using System.IO;
using Syncfusion.XlsIO;

namespace Calculated_Field
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
                IWorksheet sheet = workbook.Worksheets[1];
                IPivotTable pivotTable = sheet.PivotTables[0];

                //Add calculated field to the first pivot table
                IPivotField field = pivotTable.CalculatedFields.Add("Percent", "Units/3000*100");

                //Set Field Formula
                field.Formula = "Units/3000*200";

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CalculatedField.xlsx"));
                #endregion
            }
        }
    }
}





