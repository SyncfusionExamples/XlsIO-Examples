using Syncfusion.XlsIO;

namespace Formula_Validation
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

                //Data Validation using a custom formula
                IDataValidation validation = worksheet.CreateDataValidation(worksheet.Range["A1:A100000"]);
                validation.AllowType = ExcelDataType.Formula;
                validation.FirstFormula = "=A1>10";

                //Shows the error message
                validation.ErrorBoxText = "Enter a value greater than 10";
                validation.ErrorBoxTitle = "ERROR";
                validation.PromptBoxText = "Custom DataValidation";
                validation.ShowPromptBox = true;

                workbook.SaveAs(Path.GetFullPath("Output/FormulaValidation.xlsx"));
            }
        }
    }
}