using Microsoft.AspNetCore.Mvc;
using PivotTable.Models;
using Syncfusion.XlsIO;
using System.Diagnostics;

namespace PivotTable.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult AddPivotTable()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;

            // Open the existing Excel file using a file stream
            FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/SalesReport.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(fileStream);

            // Access the first worksheet containing the source data
            IWorksheet worksheet = workbook.Worksheets[0];

            // Create a new worksheet to hold the pivot table
            IWorksheet pivotSheet = workbook.Worksheets.Create("PivotSheet");

            // Create a pivot cache using the specified data range from the source worksheet
            IPivotCache cache = workbook.PivotCaches.Add(worksheet["A1:H50"]);

            // Create a pivot table named "PivotTable1" using the cache and place it at cell A1 in the new worksheet
            IPivotTable pivotTable = pivotSheet.PivotTables.Add("PivotTable1", pivotSheet["A1"], cache);

            // Add a column field to organize data horizontally
            pivotTable.Fields[2].Axis = PivotAxisTypes.Column;

            // Add row fields to organize data vertically
            pivotTable.Fields[3].Axis = PivotAxisTypes.Row;
            pivotTable.Fields[4].Axis = PivotAxisTypes.Row;

            // Add data fields to summarize values
            IPivotField field = pivotTable.Fields[5];
            pivotTable.DataFields.Add(field, "Units", PivotSubtotalTypes.Sum);

            field = pivotTable.Fields[6];
            pivotTable.DataFields.Add(field, "Unit Cost", PivotSubtotalTypes.Sum);

            // Apply a built-in style to the pivot table for better visual appearance
            pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleMedium14;

            // Export the workbook as a downloadable Excel file
            return ExportWorkbook(workbook, "AddPivotTable.xlsx");
        }

        public IActionResult EditPivotTable()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;

            // Open the existing Excel file using a file stream
            FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(fileStream);

            // Access the worksheet containing the pivot table
            IWorksheet worksheet = workbook.Worksheets[1];

            // Access the first pivot table in the worksheet
            IPivotTable pivotTable = worksheet.PivotTables[0];

            // Modify the pivot table structure by setting row and column fields
            pivotTable.Fields["Region"].Axis = PivotAxisTypes.Column;
            pivotTable.Fields["Employee"].Axis = PivotAxisTypes.Row;

            // Apply a new built-in style to the pivot table
            pivotTable.BuiltInStyle = PivotBuiltInStyles.PivotStyleDark2;

            // Export the modified workbook as a downloadable Excel file
            return ExportWorkbook(workbook, "EditPivotTable.xlsx");
        }

        public IActionResult RemovePivotTable()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;

            // Open the existing Excel file using a file stream
            FileStream fileStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(fileStream);

            // Access the worksheet that contains the pivot table to be removed
            IWorksheet pivotSheet = workbook.Worksheets[1];

            // Remove the pivot table named "PivotTable1" from the worksheet
            pivotSheet.PivotTables.Remove("PivotTable1");

            // Export the updated workbook as a downloadable Excel file
            return ExportWorkbook(workbook, "RemovePivotTable.xlsx");
        }

        private FileStreamResult ExportWorkbook(IWorkbook workbook, string fileName)
        {
            // Create a memory stream to hold the Excel file content
            MemoryStream stream = new MemoryStream();

            // Save the workbook to the memory stream
            workbook.SaveAs(stream);

            // Reset the stream position to the beginning
            stream.Position = 0;

            // Return the stream as a downloadable Excel file with the specified filename
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
