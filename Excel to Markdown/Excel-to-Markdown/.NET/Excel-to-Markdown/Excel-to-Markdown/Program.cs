using Syncfusion.XlsIO;

namespace Excel_to_Markdown
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(@"Data/Markdown.xlsx");

                using (FileStream fileStream = new FileStream(@"Output/ExcelToMarkdown.md", FileMode.Create, FileAccess.Write))
                {
                    workbook.SaveAs(fileStream, ExcelSaveType.Markdown);
                }
            }
        }
    }
}