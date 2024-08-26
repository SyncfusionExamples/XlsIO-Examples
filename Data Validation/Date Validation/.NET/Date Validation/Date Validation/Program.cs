using System.IO;
using Syncfusion.XlsIO;
using System;

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
                IDataValidation dateValidation = worksheet.Range["E3"].DataValidation;
                worksheet.Range["E1"].Text = "Enter the Date in E3";
                worksheet.Range["E1"].AutofitColumns();
                dateValidation.AllowType = ExcelDataType.Date;
                dateValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
                dateValidation.FirstDateTime = new DateTime(2003, 5, 10);
                dateValidation.SecondDateTime = new DateTime(2004, 5, 10);

                //Shows the error message
                dateValidation.ShowErrorBox = true;
                dateValidation.ErrorBoxText = "Enter Value between 10/5/2003 to 10/5/2004";
                dateValidation.ErrorBoxTitle = "ERROR";
                dateValidation.PromptBoxText = "Data validation for date";
                dateValidation.ShowPromptBox = true;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/DateValidation.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




