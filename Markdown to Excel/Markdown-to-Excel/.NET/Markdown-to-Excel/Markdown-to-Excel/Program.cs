using Syncfusion.XlsIO;

namespace Markdown_to_Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(@"Data/Sample.md", ExcelOpenType.Markdown);

                workbook.SaveAs(Path.GetFullPath("Output/MarkdownToExcel.xlsx"));
            }
        }
    }
}