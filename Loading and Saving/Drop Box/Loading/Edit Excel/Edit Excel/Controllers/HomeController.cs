using Dropbox.Api;
using Edit_Excel.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;
using System.Diagnostics;
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
        public async Task<IActionResult> EditDocument()
        {
            try
            {
                //Retrieve the document from DropBox
                MemoryStream stream = await GetDocumentFromDropBox();

                //Set the position as '0'
                stream.Position = 0;

                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Xlsx;

                    //Load the downloaded document
                    IWorkbook workbook = application.Workbooks.Open(stream);

                    IWorksheet worksheet = workbook.Worksheets[0];
                    worksheet.Range["A3"].Text = "Hello world";

                    //Saving the Excel to the MemoryStream 
                    MemoryStream outputStream = new MemoryStream();
                    workbook.SaveAs(outputStream);

                    //Set the position as '0'
                    outputStream.Position = 0;

                    //Download the Excel file in the browser
                    FileStreamResult fileStreamResult = new FileStreamResult(outputStream, "application/excel");
                    fileStreamResult.FileDownloadName = "EditExcel.xlsx";
                    return fileStreamResult;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return Content("Error occurred while processing the file.");
            }
        }

        /// <summary>
        /// Download file from DropBox
        /// </summary>
        public async Task<MemoryStream> GetDocumentFromDropBox()
        {
            //Define the access token for authentication with the Dropbox API
            var accessToken = "Access_Token";

            //Define the file path in Dropbox where the file is located. For ex: "/Template.docx"
            var filePathInDropbox = "FilePath";

            try
            {
                //Create a new DropboxClient instance using the provided access token
                using (var dbx = new DropboxClient(accessToken))
                {
                    //Start a download request for the specified file in Dropbox
                    using (var response = await dbx.Files.DownloadAsync(filePathInDropbox))
                    {
                        //Get the content of the downloaded file as a stream
                        var content = await response.GetContentAsStreamAsync();

                        MemoryStream stream = new MemoryStream();
                        content.CopyTo(stream);
                        return stream;
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from DropBox: {ex.Message}");
                throw; // or handle the exception as needed
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
