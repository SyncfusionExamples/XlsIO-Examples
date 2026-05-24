using Edit_Excel.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System.Diagnostics;
using Syncfusion.Drawing;

namespace Edit_Excel_file.Controllers
{
    public class HomeController : Controller
    {
        // Inject the IWebHostEnvironment service in your class or method
        private readonly IWebHostEnvironment _hostingEnvironment;
        public HomeController(IWebHostEnvironment hostingEnvironment)
        {
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult EditDocument()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Assigns default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //A existing workbook is opened.
                FileStream inputStream = new FileStream("InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Access first worksheet from the workbook.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set Text in cell A3.
                worksheet.Range["A3"].Text = "Hello World";

                //Saving the Excel to the MemoryStream 
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "EditExcel.xlsx";
                return fileStreamResult;
            }
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