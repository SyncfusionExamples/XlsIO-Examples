using Syncfusion.XlsIO;
using SkiaSharp;
using Svg.Skia;
namespace Adding_Fallback_Image_For_SVG
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Loads an svg image stream
                FileStream svgStream = new FileStream(@Path.GetFullPath(@"Data/Image.svg"), FileMode.Open);

                //Convert svg stream to png stream
                Stream pngStream = ConvertSvgStreamToPngStream(svgStream);

                //Add svg pictures in the worksheet
                worksheet.Pictures.AddPicture(1, 1, svgStream, pngStream, 400, 390);

                //Saving the workbook as stream
                FileStream stream = new FileStream("Svg.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }
        /// <summary>
        /// Convert svg stream to png stream using skiasharp
        /// </summary>
        /// <param name="svgStream"></param>
        /// <returns>"pngStream"</returns>
        public static Stream ConvertSvgStreamToPngStream(Stream svgStream)
        {
            //Create a MemoryStream to store the converted PNG
            MemoryStream pngStream = new MemoryStream();

            //Load SVG image using Skiasharp
            using (var svg = new SKSvg())
            {
                svg.Load(svgStream);

                //Get the size of SVG image
                var bitmap = new SKBitmap((int)svg.Picture.CullRect.Width, (int)svg.Picture.CullRect.Height);
                var canvas = new SKCanvas(bitmap);

                //Render SVG to the bitmap
                canvas.DrawPicture(svg.Picture);

                //Encode the bitmap as PNG and write to the MemoryStream
                var image = SKImage.FromBitmap(bitmap);
                image.Encode(SKEncodedImageFormat.Png, 100).SaveTo(pngStream);
            }
            return pngStream;
        }
    }
}
