using Syncfusion.XlsIO;

namespace Export_images_to_folder
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;

                IWorkbook workbook = application.Workbooks.Open(@"Data/Markdown.xlsx");

                MarkdownExportOptions exportOptions = new MarkdownExportOptions();
                exportOptions.SaveOptions.MarkdownExportImagesFolder = @"D:/Temp/Image1.png";

                using (FileStream fileStream = new FileStream(@"Output/Output.md", FileMode.Create, FileAccess.Write))
                {
                    workbook.SaveAs(fileStream, exportOptions);
                }
            }
        }
    }
}