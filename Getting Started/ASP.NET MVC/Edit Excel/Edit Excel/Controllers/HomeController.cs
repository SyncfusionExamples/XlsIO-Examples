using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Syncfusion.XlsIO;

namespace Edit_Excel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult EditDocument()
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Set the default application version
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load the existing Excel workbook into IWorkbook
                IWorkbook workbook = application.Workbooks.Open(Server.MapPath("App_Data/InputTemplate.xlsx"));

                //Get the first worksheet in the workbook into IWorksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign some text in a cell
                worksheet.Range["A3"].Text = "Hello World";

                //Save the Excel document
                workbook.SaveAs("EditExcel.xlsx", HttpContext.ApplicationInstance.Response, ExcelDownloadType.Open);
            }
            return View("Index");
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}