using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System.Diagnostics;
using TemplateMarker_Vertical_Argument.Models;

namespace TemplateMarker_Vertical_Argument.Controllers
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
        private static List<Employee> GetEmployeeDetails()
        {
            List<Employee> employeeList = new List<Employee>();
            Employee emp = new Employee();
            emp.Name = "Andy Bernard";
            emp.Id = 1011;
            emp.Age = 35;

            employeeList.Add(emp);

            emp = new Employee();
            emp.Name = "Jim Halpert";
            emp.Id = 1012;
            emp.Age = 26;

            employeeList.Add(emp);

            emp = new Employee();
            emp.Name = "Karen Fillippelli";
            emp.Id = 1013;
            emp.Age = 28;

            employeeList.Add(emp);

            return employeeList;
        }
        public IActionResult TemplateMarker_Vertical_Argument()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding header text
                worksheet["A1"].Text = "\"Vertical\" Argument";
                worksheet["A3"].Text = "Name";
                worksheet["B3"].Text = "Id";
                worksheet["C3"].Text = "Age";
                worksheet["A3:C3"].CellStyle.Font.Bold = true;

                //Adding markers dynamically with the argument, 'vertical'
                worksheet["A4"].Text = "%Employee.Name;vertical";
                worksheet["B4"].Text = "%Employee.Id;vertical";
                worksheet["C4"].Text = "%Employee.Age;vertical";

                //Create template marker processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("Employee", GetEmployeeDetails());

                //Apply markers
                marker.ApplyMarkers();

                //Save and close the workbook
                MemoryStream stream = new MemoryStream();
                worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "Output.xlsx";
                workbook.Close();
                excelEngine.Dispose();

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