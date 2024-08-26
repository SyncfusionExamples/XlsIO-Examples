using System.IO;
using Syncfusion.XlsIO;

namespace Text_Length_Validation
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

                //Data Validation for Text Length
                IDataValidation txtLengthValidation = worksheet.Range["A3"].DataValidation;
                worksheet.Range["A1"].Text = "Enter the Text in A3";
                worksheet.Range["A1"].AutofitColumns();
                txtLengthValidation.AllowType = ExcelDataType.TextLength;
                txtLengthValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
                txtLengthValidation.FirstFormula = "0";
                txtLengthValidation.SecondFormula = "5";

                //Shows the error message
                txtLengthValidation.ShowErrorBox = true;
                txtLengthValidation.ErrorBoxText = "Text length should be lesser than 5 characters";
                txtLengthValidation.ErrorBoxTitle = "ERROR";
                txtLengthValidation.PromptBoxText = "Data validation for text length";
                txtLengthValidation.ShowPromptBox = true;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/TextLengthValidation.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("TextLengthValidation.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
