using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using TemplateMarker_with_Insert_Argument.Models;
using System.IO.Compression;

namespace TemplateMarker_with_Insert_Argument.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
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
        public IActionResult TemplateMarker_with_Insert_Rows()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding header text
                worksheet["A1"].Text = "\"Insert\" Argument";
                worksheet["A3"].CellStyle.Font.RGBColor = Color.FromArgb(255, 0, 0);

                worksheet["A3"].Text = "\"Row\" Insertion with copy styles and copy merges";

                worksheet["A4:B4"].Merge();
                worksheet["A5:B5"].Merge();

                worksheet["A4"].Text = "Name";
                worksheet["C4"].Text = "Id";
                worksheet["D4"].Text = "Age";

                worksheet["A4:F4"].CellStyle.Font.Bold = true;

                worksheet["A5"].CellStyle.Font.Italic = true;

                //Adding markers dynamically with the arguments, 'insert','copystyles' and 'copymerges
                worksheet["A5"].Text = "%Employee.Name;insert:copystyles,copymerges";
                worksheet["C5"].Text = "%Employee.Id";
                worksheet["D5"].Text = "%Employee.Age";

                // This data will be moved to new row
                worksheet["A7"].Text = "Text in new row";

                worksheet["A9"].CellStyle.Font.RGBColor = Color.FromArgb(255, 0, 0);
                worksheet["A4:D4"].CellStyle.Color = Color.FromArgb(77, 176, 215);

                //Create template marker processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("Employee", GetEmployeeDetails());

                //Apply markers
                marker.ApplyMarkers();

                //Save and close the workbook
                MemoryStream stream = new MemoryStream();
                //worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "InsertRows.xlsx";
                workbook.Close();
                excelEngine.Dispose();

                return fileStreamResult;
            }
        }
        public IActionResult TemplateMarker_with_Insert_Columns()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding header text
                worksheet["A1"].Text = "\"Insert\" Argument";

                worksheet["A3"].CellStyle.Font.RGBColor = Color.FromArgb(255, 0, 0);
                worksheet["A4:A6"].CellStyle.Color = Color.FromArgb(77, 176, 215);

                worksheet["A3"].Text = "\"Column\" Insertion with copy styles";
                worksheet["A4"].Text = "Name";
                worksheet["A5"].Text = "Id";
                worksheet["A6"].Text = "Age";

                worksheet["A4:A6"].CellStyle.Font.Bold = true;

                worksheet["B4"].CellStyle.Color = Color.FromArgb(189, 215, 238);

                //Adding markers dynamically with the arguments, 'insert' and 'copystyles' and. 'horizontal'
                worksheet["B4"].Text = "%Employee.Name;insert:copystyles;horizontal";
                worksheet["B5"].Text = "%Employee.Id;horizontal";
                worksheet["B6"].Text = "%Employee.Age;horizontal";

                // This data will be moved to new column
                worksheet["C6"].Text = "Text in new column";

                //Create template marker processor
                ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

                //Add marker variable
                marker.AddVariable("Employee", GetEmployeeDetails());

                //Apply markers
                marker.ApplyMarkers();

                //Save and close the workbook
                MemoryStream stream = new MemoryStream();
                //worksheet.UsedRange.AutofitColumns();
                workbook.SaveAs(stream);

                //Set the position as '0'.
                stream.Position = 0;

                //Download the Excel file in the browser
                FileStreamResult fileStreamResult = new FileStreamResult(stream, "application/excel");
                fileStreamResult.FileDownloadName = "InsertColumns.xlsx";
                workbook.Close();
                excelEngine.Dispose();

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