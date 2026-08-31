using Syncfusion.XlsIO;

namespace List_Validation
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

                worksheet.Range["B1"].Text = "Item1";
                worksheet.Range["B2"].Text = "Item2";
                worksheet.Range["B3"].Text = "Item3";

                //Data validation for the user-defined range
                IDataValidation validation = worksheet.CreateDataValidation(worksheet.Range["A1:A100000"]);
                validation.AllowType = ExcelDataType.User;
                validation.FirstFormula = "=Sheet1!$B$1:$B$3";

                //Shows the error message
                validation.ErrorBoxText = "Choose the value from the list";
                validation.ErrorBoxTitle = "ERROR";
                validation.PromptBoxText = "Data validation for user-defined list";
                validation.IsPromptBoxVisible = true;
                validation.ShowPromptBox = true;

                workbook.SaveAs(Path.GetFullPath(@"Output/ListValidation.xlsx"));
            }
        }
    }
}