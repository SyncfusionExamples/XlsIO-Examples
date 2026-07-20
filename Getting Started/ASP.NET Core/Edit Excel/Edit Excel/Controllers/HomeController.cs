using Edit_Excel.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
using System.IO;

namespace Edit_Excel.Controllers
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
        public IActionResult EditDocument()
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Assigns default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //A existing workbook is opened.             
                IWorkbook workbook = application.Workbooks.Open("InputTemplate.xlsx");

                //Access first worksheet from the workbook.
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set the text in cell A3.
                worksheet.Range["A3"].Text = "Hello World";

                //Access a cell value from Excel
                var value = worksheet.Range["A1"].Value;

                //Defining the ContentType for the Excel file.
                string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                //Define the file name.
                string fileName = "EditExcel.xlsx";

                //Creating stream object.
                MemoryStream stream = new MemoryStream();

                //Saving the workbook to the stream in XLSX format.
                workbook.SaveAs(stream);

                stream.Position = 0;

                //Closing the workbook and disposing the engine.
                workbook.Close();
                excelEngine.Dispose();

                //Creates a FileStreamResult object by using the file contents, content type, and file name.
                return File(stream, ContentType, fileName);
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
