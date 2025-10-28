using System.IO;
using Syncfusion.XlsIO;

namespace Calculation_Modes
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

                //Setting calculation mode for a workbook
                workbook.CalculationOptions.CalculationMode = ExcelCalculationMode.Manual;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/CalculationMode.xlsx"));
                #endregion
            }
        }
    }
}




