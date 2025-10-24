using Syncfusion.XlsIO;
using System.IO;

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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
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

                //Saving the workbook 
                workbook.Version = ExcelVersion.Xlsx;
                workbook.SaveAs(Path.GetFullPath("Output/Replace.xlsx"));
            }
        }
    }
}




