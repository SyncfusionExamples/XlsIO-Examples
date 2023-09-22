using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System.Data;
using System.Diagnostics;
using TemplateMarker.Models;

namespace TemplateMarker.Controllers
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
        public IActionResult TemplateMarker()
        {
            //Code to read XML data to create a DataTable
            FileStream dataStream = new FileStream("Data/customers.xml", FileMode.Open, FileAccess.Read);
            DataSet customersDataSet = new DataSet();
            customersDataSet.ReadXml(dataStream, XmlReadMode.ReadSchema);
            DataTable northwindDt = customersDataSet.Tables[0];

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //load an existing file
                FileStream excelStream = new FileStream("Data/TemplateMarker.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                //Create Template Marker processor.
                //Apply the marker to export data from datatable to worksheet.
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();
                marker.AddVariable("SalesList", northwindDt);
                marker.ApplyMarkers();

                //Saving the Excel to the MemoryStream 
                MemoryStream stream = new MemoryStream();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "TemplateMarkerOutput.xlsx";
                return fileStreamResult;
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