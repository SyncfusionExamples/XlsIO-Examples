using Syncfusion.XlsIO;

namespace Date_Validation
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

                //Data Validation for Date
                IDataValidation dateValidation = worksheet.CreateDataValidation(worksheet.Range["A1:A100000"]);
                dateValidation.AllowType = ExcelDataType.Date;
                dateValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
                dateValidation.FirstDateTime = new DateTime(2003, 5, 10);
                dateValidation.SecondDateTime = new DateTime(2004, 5, 10);

                //Shows the error message
                dateValidation.ShowErrorBox = true;
                dateValidation.ErrorBoxText = "Enter a value between 10/5/2003 and 10/5/2004";
                dateValidation.ErrorBoxTitle = "ERROR";
                dateValidation.PromptBoxText = "Data validation for date";
                dateValidation.ShowPromptBox = true;

                workbook.SaveAs(Path.GetFullPath("Output/DateValidation.xlsx"));
            }
        }
    }
}