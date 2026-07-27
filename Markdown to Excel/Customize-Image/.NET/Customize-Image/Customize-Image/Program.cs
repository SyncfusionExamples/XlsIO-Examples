using Syncfusion.Office.Markdown;
using Syncfusion.XlsIO;
using System.Net;

namespace Customize_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                MdImportSettings settings = new MdImportSettings();

                settings.ImageNodeVisited += MdImportSettings_ImageNodeVisited;

                IWorkbook workbook = application.Workbooks.Open(@"Data/Sample1.md", settings);

                workbook.SaveAs(Path.GetFullPath("Output/MarkdownToExcel.xlsx"));
            }
        }
        private static void MdImportSettings_ImageNodeVisited(object sender, MdImageNodeVisitedEventArgs args)
        {
            //Set the image stream based on the image name from the input Markdown.
            if (args.Uri == "Image_1.png")
                args.ImageStream = new FileStream(Path.GetFullPath("Data/Image_1.png"), FileMode.Open);
            else if (args.Uri == "Image_2.png")
                args.ImageStream = new FileStream(Path.GetFullPath("Data/Image_2.png"), FileMode.Open);
            //Retrive the image from the website and use it.
            else if (args.Uri.StartsWith("https://"))
            {
                WebClient client = new WebClient();
                byte[] image = client.DownloadData(args.Uri);
                Stream stream = new MemoryStream(image);
                args.ImageStream = stream;
            }
        }

    }
}