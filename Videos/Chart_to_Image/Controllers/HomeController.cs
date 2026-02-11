using Chart_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System.Diagnostics;

namespace Chart_to_Image.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult ChartToPNG()
        {
            // Creates Excel engine instance
            ExcelEngine excelEngine = new ExcelEngine();

            // Access Excel application object
            IApplication application = excelEngine.Excel;

            // Sets default Excel version
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Initializes renderer
            application.XlsIORenderer = new XlsIORenderer();

            // Sets output image format as PNG
            application.XlsIORenderer.ChartRenderingOptions.ImageFormat = ExportImageFormat.Png;

            // Sets output scaling mode as Best
            application.XlsIORenderer.ChartRenderingOptions.ScalingMode = ScalingMode.Best; 

            // Opens the Excel template file
            FileStream inputStream = new FileStream("Data\\InputTemplate.xlsx", FileMode.Open, FileAccess.Read);

            // Loads workbook
            IWorkbook workbook = application.Workbooks.Open(inputStream);

            // Access worksheet sheet
            IWorksheet worksheet = workbook.Worksheets[0];

            // Reads the chart
            IChart chart = worksheet.Charts[0];

            //Returns the chart as PNG
            return ExportImage(chart, "Image.png", "image/png");
        }

        public IActionResult ChartToJPEG()
        {
            // Creates Excel engine instance
            ExcelEngine excelEngine = new ExcelEngine();

            // Access Excel application object
            IApplication application = excelEngine.Excel;
            
            // Sets default Excel version
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Initializes renderer
            application.XlsIORenderer = new XlsIORenderer();

            // Sets output image format as JPEG
            application.XlsIORenderer.ChartRenderingOptions.ImageFormat = ExportImageFormat.Jpeg;

            // Sets output scaling mode as Best
            application.XlsIORenderer.ChartRenderingOptions.ScalingMode = ScalingMode.Best;

            // Opens the Excel template file
            FileStream inputStream = new FileStream("Data\\InputTemplate.xlsx", FileMode.Open, FileAccess.Read);

            // Loads workbook
            IWorkbook workbook = application.Workbooks.Open(inputStream);

            // Access worksheet sheet
            IWorksheet worksheet = workbook.Worksheets[0];

            // Reads the chart
            IChart chart = worksheet.Charts[0];

            //Returns the chart as JPEG
            return ExportImage(chart, "Image.jpeg", "image/jpeg");
        }

        private FileStreamResult ExportImage(IChart chart, string fileName, string contentType)
        {
            // Creates memory stream for output
            MemoryStream stream = new MemoryStream();

            // Saves chart into the stream
            chart.SaveAsImage(stream);

            // Resets stream position
            stream.Position = 0;

            // Return the image as a downloadable file
            return File(stream, contentType, fileName);
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
