using Syncfusion.XlsIO;
using Syncfusion.Drawing.Fonts;
using Syncfusion.XlsIORenderer;

namespace Custom_Font
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a collection to store the custom font streams
            List<Stream> fontStreams = new List<Stream>();

            // Retrieve all font files from the specified directory
            foreach (string file in Directory.GetFiles(Path.GetFullPath(@"Data/MyFonts")))
            {
                string extension = Path.GetExtension(file);

                // Load only TrueType and OpenType font files
                if (extension.Equals(".ttf", StringComparison.OrdinalIgnoreCase) ||
                    extension.Equals(".otf", StringComparison.OrdinalIgnoreCase))
                {
                    // Read the font file and copy its content to a memory stream
                    FileStream fileStream = new FileStream(file, FileMode.OpenOrCreate, FileAccess.Read);
                    Stream stream = new MemoryStream();
                    fileStream.CopyTo(stream);
                    stream.Position = 0;
                    fontStreams.Add(stream);
                }
            }

            // Register the custom fonts for use during Excel-to-Image conversion
            FontManager.RegisterFonts(fontStreams);

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet sheet = workbook.Worksheets[0];

                application.XlsIORenderer = new XlsIORenderer();

                using (FileStream outputStream = new FileStream(Path.GetFullPath("Output/Image.png"), FileMode.Create, FileAccess.Write))
                {
                    sheet.ConvertToImage(sheet.UsedRange, outputStream);
                }
            }

            // Clear the registered fonts and dispose their associated streams
            FontManager.ClearRegisteredFonts(true);
        }
    }
}