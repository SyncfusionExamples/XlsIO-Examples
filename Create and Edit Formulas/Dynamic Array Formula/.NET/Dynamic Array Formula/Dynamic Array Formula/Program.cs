using Syncfusion.XlsIO;

namespace Dynamic_Array_Formula
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
                IWorksheet worksheet = workbook.Worksheets[0];

                // Source data for SORT, FILTER, and UNIQUE.
                worksheet["A1"].Text = "Numbers";
                worksheet["A2"].Number = 30;
                worksheet["A3"].Number = 10;
                worksheet["A4"].Number = 20;
                worksheet["A5"].Number = 10;
                worksheet["A6"].Number = 40;

                // 1. Generates numbers from 1 to 5.
                worksheet["C1"].Text = "SEQUENCE";
                worksheet["C2"].FormulaArray = "=SEQUENCE(5)";

                // 2. Sorts the numbers in ascending order.
                worksheet["E1"].Text = "SORT";
                worksheet["E2"].FormulaArray = "=SORT(A2:A6)";

                // 3. Returns only numbers greater than 20.
                worksheet["G1"].Text = "FILTER";
                worksheet["G2"].FormulaArray = "=FILTER(A2:A6,A2:A6>20)";

                // 4. Returns the unique numbers.
                worksheet["I1"].Text = "UNIQUE";
                worksheet["I2"].FormulaArray = "=UNIQUE(A2:A6)";

                // 5. Generates five random numbers.
                worksheet["K1"].Text = "RANDARRAY";
                worksheet["K2"].FormulaArray = "=RANDARRAY(5,1,1,10,TRUE)";

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Formula.xlsx"));
                #endregion
            }
        }
    }
}