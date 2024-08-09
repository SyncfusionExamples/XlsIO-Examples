using Syncfusion.XlsIO;

namespace Find
{
    class Program
    {
        public static void Main(string[] args) 
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream fileStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Searches for the given string within the text of worksheet
                IRange[] result1 = worksheet.FindAll("Gill", ExcelFindType.Text);

                //Searches for the given string within the text of worksheet
                IRange[] result2 = worksheet.FindAll(700, ExcelFindType.Number);

                //Searches for the given string in formulas
                IRange[] result3 = worksheet.FindAll("=SUM(F10:F11)", ExcelFindType.Formula);

                //Searches for the given string in calculated value, number and text
                IRange[] result4 = worksheet.FindAll("41", ExcelFindType.Values);

                //Searches for the given string in comments
                IRange[] result5 = worksheet.FindAll("Desk", ExcelFindType.Comments);

                //Searches for the given string within the text of worksheet and case matched
                IRange[] result6 = worksheet.FindAll("Pen Set", ExcelFindType.Text, ExcelFindOptions.MatchCase);

                //Searches for the given string within the text of worksheet and the entire cell content matching to search text
                IRange[] result7 = worksheet.FindAll("5", ExcelFindType.Text, ExcelFindOptions.MatchEntireCellContent);

                //Saving the workbook as stream
                FileStream stream = new FileStream("Find.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Find.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}