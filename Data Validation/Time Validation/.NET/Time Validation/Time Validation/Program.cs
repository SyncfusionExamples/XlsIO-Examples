﻿using System.IO;
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
                IDataValidation timeValidation = worksheet.Range["B3"].DataValidation;
                worksheet.Range["B1"].Text = "Enter the time between 10:00 and 12:00 'o Clock in B3";
                worksheet.Range["B1"].AutofitColumns();
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

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/TimeValidation.xlsx"));
                #endregion
            }
        }
    }
}




