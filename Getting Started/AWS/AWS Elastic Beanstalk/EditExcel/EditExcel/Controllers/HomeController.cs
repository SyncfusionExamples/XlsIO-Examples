using EditExcel.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;
using System.Diagnostics;
using static Syncfusion.XlsIO.Parser.Biff_Records.Charts.ChartPicfRecord;

namespace Create_Excel.Controllers
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

        public ActionResult EditExcelDocument()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Assigns default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //A existing workbook is opened.             
                FileStream sampleFile = new FileStream("Data/InputTemplate.xlsx", FileMode.Open);
                IWorkbook workbook = application.Workbooks.Open(sampleFile);

                //Access first worksheet from the workbook.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set Text in cell A3.
                worksheet.Range["A3"].Text = "Hello World";
                //Saving the Excel to the MemoryStream 
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download Excel document in the browser.
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Sample.xlsx");
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
