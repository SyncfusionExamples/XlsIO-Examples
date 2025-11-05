using Convert_Excel_to_Image.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Convert_Excel_to_Image.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult CreateDocument()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open("Sample.xlsx");
                IWorksheet worksheet = workbook.Worksheets[0];

                //Initialize XlsIO renderer.
                application.XlsIORenderer = new XlsIORenderer();

                //Create the MemoryStream to save the image.      
                MemoryStream imageStream = new MemoryStream();

                //Save the converted image to MemoryStream.
                worksheet.ConvertToImage(worksheet.UsedRange, imageStream);
                imageStream.Position = 0;

                //Download image in the browser.
                return File(imageStream, "application/jpeg", "Sample.jpeg");
            }
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