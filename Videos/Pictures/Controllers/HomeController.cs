using Microsoft.AspNetCore.Mvc;
using Pictures.Models;
using Syncfusion.XlsIO;
using System.Diagnostics;

namespace Pictures.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult AddPicture()
        {
            // Create a new Excel engine instance
            ExcelEngine excelEngine = new ExcelEngine();

            // Access the Excel application
            IApplication application = excelEngine.Excel;

            // Set default Excel version to XLSX
            application.DefaultVersion = ExcelVersion.Xlsx;

           // Open the input workbook stream
            FileStream fileStream = new FileStream("Data\\Input.xlsx", FileMode.Open, FileAccess.Read);

            // Open the workbook from the file stream
            IWorkbook workbook = application.Workbooks.Open(fileStream);

            // Get the first worksheet
            IWorksheet worksheet = workbook.Worksheets[0];

            // Open the image stream to be inserted
            FileStream imageStream = new FileStream("Data\\Image.png", FileMode.Open, FileAccess.Read);

            // Add the image at row 10, column 2
            IPictureShape picture = worksheet.Pictures.AddPicture(10, 2, imageStream);

            // Set the picture width
            picture.Width = 320;

            // Set the picture height
            picture.Height = 180;

            //Return the workbook as a downloadable file
            return ExportWorkbook(workbook, "AddPicture.xlsx");
        }

        public IActionResult ModifyPicture()
        {
            // Create a new Excel engine instance
            ExcelEngine excelEngine = new ExcelEngine();

            // Access the Excel application
            IApplication application = excelEngine.Excel;

            // Set default Excel version to XLSX
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the input workbook stream
            FileStream fileStream = new FileStream("Data\\Template.xlsx", FileMode.Open, FileAccess.Read);

            // Open the workbook from the stream
            IWorkbook workbook = application.Workbooks.Open(fileStream);

            // Get the first worksheet
            IWorksheet worksheet = workbook.Worksheets[0];

            // Get the first picture in the worksh
            IPictureShape picture = worksheet.Pictures[0];

            // Set vertical position 
            picture.Top = 150;

            // Set horizontal position 
            picture.Left = 200;

            // Set the new width of the picture
            picture.Width = 420;

            // Set the new height of the picture
            picture.Height = 240;

            //Return the workbook as a downloadable file
            return ExportWorkbook(workbook, "ModifyPicture.xlsx");
        }

        public IActionResult RemovePicture()
        {
            // Create a new Excel engine instance
            ExcelEngine excelEngine = new ExcelEngine();

            // Access the Excel application
            IApplication application = excelEngine.Excel;

            // Set default Excel version to XLSX
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the template workbook stream
            FileStream fileStream = new FileStream("Data\\Template.xlsx", FileMode.Open, FileAccess.Read);

            // Open the workbook from the stream
            IWorkbook workbook = application.Workbooks.Open(fileStream);

            // Get the first worksheet
            IWorksheet worksheet = workbook.Worksheets[0];

            // Get the first picture in the worksheet
            IPictureShape picture = worksheet.Pictures[0];

            // Remove the selected picture from the worksheet
            picture.Remove();


           // return the updated workbook as a downloadable file
            return ExportWorkbook(workbook, "RemovePicture.xlsx");
        }

        private FileStreamResult ExportWorkbook(IWorkbook workbook, string fileName)
        {
            // Create an in-memory stream to hold the workbook
            MemoryStream stream = new MemoryStream();

            // Save the workbook into the memory stream
            workbook.SaveAs(stream);

            // Reset the stream position to the beginnin
            stream.Position = 0;

            // Return the stream as a downloadable file
            return File(stream, "application/xlsx", fileName);
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
