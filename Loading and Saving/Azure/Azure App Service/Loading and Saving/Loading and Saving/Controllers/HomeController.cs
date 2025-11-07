using System.Diagnostics;
using Loading_and_Saving.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;

namespace Loading_and_Saving.Controllers
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
        public IActionResult LoadingandSaving()
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing Excel document
                IWorkbook workbook = application.Workbooks.Open("Data/InputTemplate.xlsx");

                //Access first worksheet from the workbook.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set Text in cell A3.
                worksheet.Range["A3"].Text = "Hello World";

                //Save the Excel to MemoryStream 
                MemoryStream outputStream = new MemoryStream();
                workbook.SaveAs(outputStream);

                //Set the position
                outputStream.Position = 0;

                //Download the Excel document in the browser.
                return File(outputStream, "application/msexcel", "Output.xlsx");
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
