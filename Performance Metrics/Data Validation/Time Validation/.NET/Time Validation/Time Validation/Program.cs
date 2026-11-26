using Syncfusion.XlsIO;

namespace Time_Validation
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

                //Data validation for the time
                IDataValidation timeValidation = worksheet.CreateDataValidation(worksheet.Range["A1:A100000"]);
                timeValidation.AllowType = ExcelDataType.Time;
                timeValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
                timeValidation.FirstFormula = "10.00";
                timeValidation.SecondFormula = "12.00";

                //Shows the error message
                timeValidation.ShowErrorBox = true;
                timeValidation.ErrorBoxText = "Enter a correct time";
                timeValidation.ErrorBoxTitle = "ERROR";
                timeValidation.PromptBoxText = "Data validation for time";
                timeValidation.ShowPromptBox = true;

                workbook.SaveAs(Path.GetFullPath("Output/TimeValidation.xlsx"));
            }
        }
    }
}




