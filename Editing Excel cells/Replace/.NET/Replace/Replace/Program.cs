using Syncfusion.XlsIO;

namespace Replace
{
    class program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Replaces the given string with another string
                worksheet.Replace("Wilson", "William");

                //Replaces the given string with another string on match case
                worksheet.Replace("4.99", "4.90", ExcelFindOptions.MatchCase);

                //Replaces the given string with another string matching entire cell content to the search word
                worksheet.Replace("Pen Set", "Pen", ExcelFindOptions.MatchEntireCellContent);

                //Replaces the given string with DateTime value
                worksheet.Replace("DateValue",DateTime.Now);

                //Replaces the given string with Array
                worksheet.Replace("Central", new string[] { "Central", "East" }, true);

                //Saving the workbook as stream
                FileStream stream = new FileStream("Replace.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.Version = ExcelVersion.Xlsx;
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }
    }
}
