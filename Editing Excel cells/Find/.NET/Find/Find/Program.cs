using Syncfusion.XlsIO;
using System.IO;

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
                FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Searches for the given string within the text of worksheet
                IRange[] textCells = worksheet.FindAll("Gill", ExcelFindType.Text);

                //Searches for the given string within the text of worksheet
                IRange[] numberCells = worksheet.FindAll(700, ExcelFindType.Number);

                //Searches for the given string in formulas
                IRange[] formulaCells = worksheet.FindAll("=SUM(F10:F11)", ExcelFindType.Formula);

                //Searches for the given string in calculated value, number and text
                IRange[] valueCells = worksheet.FindAll("41", ExcelFindType.Values);

                //Searches for the given string within the text of worksheet and case matched
                IRange[] textMatchingCaseCells = worksheet.FindAll("Pen Set", ExcelFindType.Text, ExcelFindOptions.MatchCase);

                //Searches for the given string within the text of worksheet and the entire cell content matching to search text
                IRange[] textMatchingEntireContentCells = worksheet.FindAll("5", ExcelFindType.Text, ExcelFindOptions.MatchEntireCellContent);

                foreach (IRange cell in textCells)
                {
                    //Highlight found text cells in red
                    cell.CellStyle.Color = Syncfusion.Drawing.Color.FromArgb(255, 255, 0, 0); 
                }

                foreach (IRange cell in numberCells)
                {
                    //Highlight found number cells in green
                    cell.CellStyle.Color = Syncfusion.Drawing.Color.FromArgb(255, 0, 255, 0); 
                }

                foreach (IRange cell in formulaCells)
                {
                    //Highlight found formula cells in blue
                    cell.CellStyle.Color = Syncfusion.Drawing.Color.FromArgb(255, 0, 0, 255); 
                }

                foreach (IRange cell in valueCells)
                {
                    //Highlight found value cells in orange
                    cell.CellStyle.Color = Syncfusion.Drawing.Color.FromArgb(255, 255, 165, 0); 
                }

                foreach (IRange cell in textMatchingCaseCells)
                {
                    //Highlight found case-matching text cells in purple
                    cell.CellStyle.Color = Syncfusion.Drawing.Color.FromArgb(255, 128, 0, 128); 
                }

                foreach (IRange cell in textMatchingEntireContentCells)
                {
                    //Highlight found entire content matching text cells in teal
                    cell.CellStyle.Color = Syncfusion.Drawing.Color.FromArgb(255, 0, 128, 128); 
                }

                //Saving the workbook as stream
                FileStream stream = new FileStream(Path.GetFullPath(@"Output/Find.xlsx"), FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }
    }
}




