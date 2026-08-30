using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Excel_to_Markdown_Shapes
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                application.XlsIORenderer = new XlsIORenderer();

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                MarkdownExportOptions markdownExportOptions = new MarkdownExportOptions();
                markdownExportOptions.PreserveEmptyRow = true;

                workbook.SaveAs(Path.GetFullPath(@"Output/ExcelToMarkdown.md"), markdownExportOptions);
            }
        }
    }
}