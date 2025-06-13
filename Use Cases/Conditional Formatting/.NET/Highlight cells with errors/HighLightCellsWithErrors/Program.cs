using Syncfusion.XlsIO;

class Program
{ 
    public static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            //Add some data and errors
            worksheet.Range["A1"].Text = "Sample Data";

            //Creates a #DIV/0! error
            worksheet.Range["A2"].Formula = "=1/0";

            //Creates a #N/A error
            worksheet.Range["A3"].Formula = "=VLOOKUP(\"NonExistent\",B1:C5,2,FALSE)"; 

            //Define the range to apply formatting
            IRange range = worksheet.Range["A1:A10"];

            //Add conditional formatting to highlight cells with errors
            IConditionalFormats conditionalFormats = range.ConditionalFormats;
            IConditionalFormat conditionalFormat = conditionalFormats.AddCondition();

            //Set format type to ContainsErrors
            conditionalFormat.FormatType = ExcelCFType.ContainsErrors;

            //Apply red background to cells containing errors
            conditionalFormat.BackColor = ExcelKnownColors.Red;

            #region Save
            //Saving the workbook
            FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);
            #endregion

            //Dispose streams
            outputStream.Dispose();
        }
    }
}