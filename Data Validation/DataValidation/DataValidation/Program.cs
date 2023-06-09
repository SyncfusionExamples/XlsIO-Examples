
using Syncfusion.XlsIO;
using static System.Net.Mime.MediaTypeNames;
using (ExcelEngine excelEngine = new ExcelEngine())
{
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet sheet = workbook.Worksheets[0];

    //Adding list validation
    IDataValidation listValidation = sheet.Range["C7"].DataValidation;
    sheet.Range["B7"].Text = "Select an item from the dropdown list";
    listValidation.ListOfValues = new string[] { "Brand", "Price", "Product" };
    listValidation.PromptBoxText = "List validation";
    listValidation.IsPromptBoxVisible = true;
    listValidation.ShowPromptBox = true;

    //Adding number validation
    IDataValidation numbervalidation = sheet.Range["C9"].DataValidation;
    sheet.Range["B9"].Text = "Enter a number between 0 to 10";
    numbervalidation.AllowType = ExcelDataType.Integer;
    numbervalidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
    numbervalidation.FirstFormula = "0";
    numbervalidation.SecondFormula = "10";
    numbervalidation.ShowErrorBox = true;
    numbervalidation.ErrorBoxText = "Enter value between only 0 to 10";
    numbervalidation.ErrorBoxTitle = "ERROR";
    numbervalidation.PromptBoxText = "Number validation";
    numbervalidation.ShowPromptBox = true;

    //Adding date validation
    IDataValidation dateValidation = sheet.Range["C11"].DataValidation;
    sheet.Range["B11"].Text = "Enter a date between 5/10/2003 to 5/10/2004";
    dateValidation.AllowType = ExcelDataType.Date;
    dateValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
    dateValidation.FirstDateTime = new DateTime(2003, 5, 10);
    dateValidation.SecondDateTime = new DateTime(2004, 5, 10);
    dateValidation.ShowErrorBox = true;
    dateValidation.ErrorBoxText = "Enter value between 5/10/2003 to 5/10/2004";
    dateValidation.ErrorBoxTitle = "ERROR";
    dateValidation.PromptBoxText = "Date validation";
    dateValidation.ShowPromptBox = true;

    //Adding text length validation
    IDataValidation textValidation = sheet.Range["C13"].DataValidation;
    sheet.Range["B13"].Text = "Enter a text of 6 characters or less";
    textValidation.AllowType = ExcelDataType.TextLength;
    textValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
    textValidation.FirstFormula = "1";
    textValidation.SecondFormula = "6";
    textValidation.ShowErrorBox = true;
    textValidation.ErrorBoxText = "Enter a text with length of maximum 6 characters";
    textValidation.ErrorBoxTitle = "ERROR";
    textValidation.PromptBoxText = "Text length validation";
    textValidation.ShowPromptBox = true;

    //Adding time validation
    IDataValidation timeValidation = sheet.Range["C15"].DataValidation;
    sheet.Range["B15"].Text = "Enter a time between 10:00 AM to 12:00 PM";
    timeValidation.AllowType = ExcelDataType.Time;
    timeValidation.CompareOperator = ExcelDataValidationComparisonOperator.Between;
    timeValidation.FirstFormula = "10:00";
    timeValidation.SecondFormula = "12:00";
    timeValidation.ShowErrorBox = true;
    timeValidation.ErrorBoxText = "Enter the time between 10 to 12 ";
    timeValidation.ErrorBoxTitle = "ERROR";
    timeValidation.PromptBoxText = "Time validation";
    timeValidation.ShowPromptBox = true;


    //Adding time validation
    IDataValidation formulaValidation = sheet.Range["C17"].DataValidation;
    sheet.Range["B17"].Text = "Enter a negative number";
    formulaValidation.AllowType = ExcelDataType.Formula;
    formulaValidation.FirstFormula = "=C17 < 0";
    formulaValidation.ShowErrorBox = true;
    formulaValidation.ErrorBoxText = "Enter only negative numbers";
    formulaValidation.ErrorBoxTitle = "ERROR";
    formulaValidation.PromptBoxText = "Formula validation";
    formulaValidation.ShowPromptBox = true;

    sheet.Range["B2:C2"].Merge();

    sheet.Range["B2"].Text = "Data validation";
    sheet.Range["B5"].Text = "Validation criteria";
    sheet.Range["C5"].Text = "Validation";
    sheet.Range["B5"].CellStyle.Font.Bold = true;
    sheet.Range["C5"].CellStyle.Font.Bold = true;
    sheet.Range["B2"].CellStyle.Font.Bold = true;
    sheet.Range["B2"].CellStyle.Font.Size = 16;
    sheet.Range["B2"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

    sheet.UsedRange.AutofitColumns();
    sheet.UsedRange.AutofitRows();

    //Saving the workbook
    FileStream outputStream = new FileStream("DataValidation.xlsx", FileMode.Create, FileAccess.Write);
    workbook.SaveAs(outputStream);
    outputStream.Dispose();
}