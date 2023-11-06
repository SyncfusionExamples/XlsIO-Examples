using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.XlsIO;
using TemplateMarker_Horizontal_Argument.Models;

namespace TemplateMarker_Horizontal_Argument.Controllers
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
        public IActionResult TemplateMarker_Horizontal_Argument()
        {
            
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding header text
                worksheet["A1"].Text = "\"Horizontal\" Argument";
                worksheet["A3"].Text = "Name";
                worksheet["A4"].Text = "Id";
                worksheet["A5"].Text = "Age";
                worksheet["A3:A5"].CellStyle.Font.Bold = true;

                //Adding markers dynamically with the argument, 'horizontal'
                worksheet["B3"].Text = "%Employee.Name;horizontal";
                worksheet["B4"].Text = "%Employee.Id;horizontal";
                worksheet["B5"].Text = "%Employee.Age;horizontal";

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