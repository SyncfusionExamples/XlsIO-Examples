using Syncfusion.Office.Markdown;
using Syncfusion.XlsIO;

namespace Customize_image_path
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
                exportOptions.SaveOptions.ImageNodeVisited += MdExportSettings_ImageNodeVisited;

                using (FileStream fileStream = new FileStream(@"Output/Output.md", FileMode.Create, FileAccess.Write))
                {
                    workbook.SaveAs(fileStream, exportOptions);
                }
            }
        }

        private static void MdExportSettings_ImageNodeVisited(object sender, SaveImageNodeVisitedEventArgs args)
        {
            string imagepath = @"D:\Temp\Image1.png";
            //Save the image stream as a file. 
            using (FileStream fileStreamOutput = File.Create(imagepath))
                args.ImageStream.CopyTo(fileStreamOutput);
            //Set the image URI to be used in the output markdown.
            args.Uri = imagepath;
        }
    }
}