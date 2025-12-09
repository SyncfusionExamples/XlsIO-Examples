using System.Diagnostics;
using Comments.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;

namespace Comments.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult AddComment()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the Excel workbook from the specified template
            IWorkbook workbook = application.Workbooks.Open("Data\\CommentsTemplate.xlsx", ExcelOpenType.Automatic);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Add a threaded comment to cell H16
            IThreadedComment threadedComment = worksheet.Range["H16"].AddThreadedComment(
                "What is the reason for the higher total amount of \"desk\" in the west region?",
                "User1",
                DateTime.Now
            );

            // Export the modified workbook as a downloadable Excel file
            return ExportWorkbook(workbook, "AddComment.xlsx");
        }

        public IActionResult ReplyComment()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the workbook containing the comment to reply to
            IWorkbook workbook = application.Workbooks.Open("Data\\ReplyInput.xlsx", ExcelOpenType.Automatic);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Access the collection of threaded comments
            IThreadedComments threadedComments = worksheet.ThreadedComments;

            // Add a reply to the first threaded comment
            threadedComments[0].AddReply(
                "The unit cost of desk is higher compared to other items in the west region. As a result, the total amount is elevated.",
                "User2",
                DateTime.Now
            );

            // Export the modified workbook
            return ExportWorkbook(workbook, "ReplyComment.xlsx");
        }

        public IActionResult ResolveComment()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the workbook containing the comment to resolve
            IWorkbook workbook = application.Workbooks.Open("Data\\ResolveInput.xlsx", ExcelOpenType.Automatic);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Access the threaded comments
            IThreadedComments threadedComments = worksheet.ThreadedComments;

            // Mark the first comment as resolved
            threadedComments[0].IsResolved = true;

            // Export the updated workbook
            return ExportWorkbook(workbook, "ResolveComment.xlsx");
        }

        public IActionResult DeleteComment()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the workbook containing the comment to delete
            IWorkbook workbook = application.Workbooks.Open("Data\\DeleteInput.xlsx", ExcelOpenType.Automatic);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Access the threaded comments
            IThreadedComments threadedComments = worksheet.ThreadedComments;

            // Delete the first threaded comment
            threadedComments[0].Delete();

            // Export the updated workbook
            return ExportWorkbook(workbook, "DeleteComment.xlsx");
        }

        public IActionResult ClearComment()
        {
            // Initialize the Excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open the workbook containing comments to clear
            IWorkbook workbook = application.Workbooks.Open("Data\\ClearInput.xlsx", ExcelOpenType.Automatic);
            IWorksheet worksheet = workbook.Worksheets[0];

            // Access all threaded comments
            IThreadedComments threadedComments = worksheet.ThreadedComments;

            // Clear all threaded comments from the worksheet
            threadedComments.Clear();

            // Export the cleaned workbook
            return ExportWorkbook(workbook, "ClearComment.xlsx");
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
