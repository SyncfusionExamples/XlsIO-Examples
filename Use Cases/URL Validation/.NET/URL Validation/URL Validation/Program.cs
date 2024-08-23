using Syncfusion.XlsIO;

namespace URL_Validation
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                // Data validation for url format
                IDataValidation validation = worksheet.Range["G2:G7"].DataValidation;
                validation.AllowType = ExcelDataType.Formula;
                validation.FirstFormula = "=AND(ISNUMBER(SEARCH(\"://\", G2:G7)), ISNUMBER(SEARCH(\".\", G2:G7, SEARCH(\"://\", G2:G7))))";

                //Shows the error message
                validation.ErrorBoxText = "Please enter a valid URL.";
                validation.ErrorBoxTitle = "Invalid URL Format";
                validation.PromptBoxText = "Enter a valid URL";
                validation.IsPromptBoxVisible = true;
                validation.ShowPromptBox = true;

                //Saving the workbook as stream
                FileStream OutputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(OutputStream);

                //Dispose stream
                inputStream.Dispose();
                OutputStream.Dispose();
            }
        }
    }
}
