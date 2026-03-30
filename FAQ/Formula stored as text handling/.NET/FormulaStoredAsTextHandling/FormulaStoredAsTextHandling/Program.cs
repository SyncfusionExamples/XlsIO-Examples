using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using static Syncfusion.XlsIO.Implementation.WorksheetImpl;

namespace FormulaStoredAsTextHandling
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Create a new workbook with one worksheet
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign a formula as text to the cell
                worksheet["A1"].Text = "=SUM(2+2)";

                //Get the cell type of
                Syncfusion.XlsIO.Implementation.WorksheetImpl.TRangeValueType cellType =
                    (worksheet as WorksheetImpl).GetCellType(1, 1, false);

                //Check if the cell type is string
                if (cellType == TRangeValueType.String)
                {
                    //Retrieve the formula text
                    string formulaText = worksheet["A1"].Text;

                    //Clear the cell value
                    worksheet["A1"].Value = string.Empty;

                    //Reassign the formula to the cell
                    worksheet["A1"].Formula = formulaText;
                }

                //Save the workbook
                workbook.SaveAs("../../../Output/Output.xlsx");
            }
        }
    }
}