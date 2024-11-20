using Syncfusion.XlsIO;

namespace Custom_Validation
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Initialize Excel engine and application.
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                // Create a workbook and worksheet.
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                // Data validation for custom data.
                IDataValidation validation = worksheet.Range["A3"].DataValidation;
                worksheet.Range["A1"].Text = "Enter the value greater than 10 in A1";
                worksheet.Range["A2"].Text = "Enter the text in A3";
                worksheet.Range["A1"].AutofitColumns();
                validation.AllowType = ExcelDataType.Formula;
                validation.FirstFormula = "=A1>10";

                // Show the error message.
                validation.ShowErrorBox = true;
                validation.ErrorBoxText = "A1 value is less than 10";
                validation.ErrorBoxTitle = "ERROR";
                validation.PromptBoxText = "Custom Data Validation";
                validation.ShowPromptBox = true;

                // Save the Excel document.
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/CustomValidation.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}