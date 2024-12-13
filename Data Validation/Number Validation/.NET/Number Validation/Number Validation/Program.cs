using System.IO;
using Syncfusion.XlsIO;

namespace Number_Validation
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

                //Data Validation for Numbers
                IDataValidation numberValidation = worksheet.Range["D3"].DataValidation;
                worksheet.Range["D1"].Text = "Enter the Number in D3";
                worksheet.Range["D1"].AutofitColumns();
                numberValidation.AllowType = ExcelDataType.Integer;
                numberValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
                numberValidation.FirstFormula = "0";
                numberValidation.SecondFormula = "10";

                //Shows the error message
                numberValidation.ShowErrorBox = true;
                numberValidation.ErrorBoxText = "Enter a value between 0 and 10";
                numberValidation.ErrorBoxTitle = "ERROR";
                numberValidation.PromptBoxText = "Data validation for numbers";
                numberValidation.ShowPromptBox = true;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/NumberValidation.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




