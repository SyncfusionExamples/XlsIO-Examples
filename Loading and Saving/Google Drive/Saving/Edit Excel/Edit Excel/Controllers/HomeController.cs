using Edit_Excel.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.XlsIO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;

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
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Download the document from Google Drive
                MemoryStream stream = await GetDocumentFromGoogleDrive();

                //Set the position as '0'
                stream.Position = 0;

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
        /// <summary>
        /// Download file from Google Drive
        /// </summary>
        public async Task<MemoryStream> GetDocumentFromGoogleDrive()
        {
            //Define the path to the service account key file
            string serviceAccountKeyPath = "Your_service_account_key_path";

            //Specify the FileID of the file to download
            string fileID = "Your_file_id"; 

            try
            {
                //Authenticate the Google Drive API access using the service account key
                GoogleCredential credential = GoogleCredential.FromFile(serviceAccountKeyPath).CreateScoped(DriveService.ScopeConstants.Drive);

                //Create the Google Drive service
                DriveService service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential
                });

                //Create a request to get the file from Google Drive
                var request = service.Files.Get(fileID);

                //Download the file into a MemoryStream
                MemoryStream stream = new MemoryStream();
                await request.DownloadAsync(stream);

                return stream;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving document from Google Drive: {ex.Message}");
                throw;
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
