using System.IO;
using Syncfusion.XlsIO;

namespace UserDefinedValidation
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
                IWorksheet worksheet = workbook.Worksheets[0];

                //Data validation for the user-defined range
                IDataValidation validation = worksheet.Range["C3"].DataValidation;
                validation.AllowType = ExcelDataType.User;
                validation.FirstFormula = "=Sheet1!$B$1:$B$3";
                worksheet.Range["C1"].Text = "Data Validation List in C3";
                worksheet.Range["C1"].AutofitColumns();
               
                //Shows the error message
                validation.ErrorBoxText = "Choose the value from the list";
                validation.ErrorBoxTitle = "ERROR";
                validation.PromptBoxText = "Data validation for user-defined list";
                validation.IsPromptBoxVisible = true;
                validation.ShowPromptBox = true;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath(@"Output/ListValidation.xlsx"));
                #endregion
            }
        }
    }
}