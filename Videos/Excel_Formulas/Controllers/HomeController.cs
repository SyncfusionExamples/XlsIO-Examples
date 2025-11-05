using System.Diagnostics;
using Excel_Formulas.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;

namespace Excel_Formulas.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult AddFormula()
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Number = 10;
            worksheet.Range["B1"].Number = 20;

            worksheet.Range["C1"].Formula = "=SUM(A1,B1)";

            return ExportWorkbook(workbook, "Formula.xlsx");
        }

        public IActionResult CrossSheetReference()
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(2);
            IWorksheet worksheet1 = workbook.Worksheets[0];
            IWorksheet worksheet2 = workbook.Worksheets[1];

            worksheet1.Range["A1"].Number = 10;
            worksheet2.Range["B1"].Number = 20;

            worksheet1.Range["C1"].Formula = "=SUM(Sheet2!B1,Sheet1!A1)";

            return ExportWorkbook(workbook, "CrossSheetReference.xlsx");
        }

        public IActionResult NamedRanges()
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1"].Number = 10;
            worksheet.Range["B1"].Number = 20;

            IName name1 = workbook.Names.Add("One");
            name1.RefersToRange = worksheet.Range["A1"];

            IName name2 = workbook.Names.Add("Two");
            name2.RefersToRange = worksheet.Range["B1"];

            worksheet.Range["C1"].Formula = "=SUM(One,Two)";

            return ExportWorkbook(workbook, "NamedRanges.xlsx");
        }

        public IActionResult ArrayFormula()
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["A1:D1"].FormulaArray = "{1,2,3,4}";

            worksheet.Names.Add("ArrayRange", worksheet.Range["A1:D1"]);

            worksheet.Range["A2:D2"].FormulaArray = "ArrayRange+100";

            return ExportWorkbook(workbook, "ArrayFormula.xlsx");
        }

        public IActionResult TableFormula()
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;
            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet worksheet = workbook.Worksheets[0];

            IListObject table = worksheet.ListObjects.Create("Table1", worksheet["A1:D3"]);

            worksheet[1, 1].Text = "Products";
            worksheet[1, 2].Text = "Rate";
            worksheet[1, 3].Text = "Quantity";
            worksheet[1, 4].Text = "Total";

            worksheet[2, 1].Text = "Item1";
            worksheet[2, 2].Number = 200;
            worksheet[2, 3].Number = 2;

            worksheet[3, 1].Text = "Item2";
            worksheet[3, 2].Number = 300;
            worksheet[3, 3].Number = 3;

            table.Columns[3].CalculatedFormula = "SUM(20,[Rate]*[Quantity])";

            return ExportWorkbook(workbook, "CalculatedColumn.xlsx");
        }

        private FileStreamResult ExportWorkbook(IWorkbook workbook, string fileName)
        {
            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
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
