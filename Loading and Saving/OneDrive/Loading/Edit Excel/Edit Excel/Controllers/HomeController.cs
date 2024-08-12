using Edit_Excel.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Identity.Client;
using Syncfusion.XlsIO;
using System.Diagnostics;
using System.Net.Http.Headers;

namespace Edit_Excel.Controllers
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
        public async Task<IActionResult> EditDocument()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Download the document from OneDrive
                MemoryStream stream = await DownloadDocumentFromOneDrive();

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
        /// Download file from OneDrive
        /// </summary>
        public async Task<MemoryStream> DownloadDocumentFromOneDrive()
        {
            //Replace with your application (client) ID, tenant ID, and secret
            string clientId = "your-client-id";
            string tenantId = "your-tenant-id";
            string clientSecret = "your-client-secret";

            //Replace with the user ID (email address) whose OneDrive you want to access
            string userId = "user@example.com";

            //Replace with the OneDrive file path where you want to download the file For ex: "/Template.xlsx"
            string filePath = "FilePath";

            //Initialize the MSAL client
            var confidentialClientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithClientSecret(clientSecret)
                .WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
                .Build();

            //Acquire an access token
            string[] scopes = { "https://graph.microsoft.com/.default" };
            var authenticationResult = await confidentialClientApplication
                .AcquireTokenForClient(scopes)
                .ExecuteAsync();

            //Create an HTTP client with the access token
            var httpClient = new HttpClient();
            httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", authenticationResult.AccessToken);

            //Construct the OneDrive download URL using user ID and file path
            var downloadUrl = $"https://graph.microsoft.com/v1.0/users/{userId}/drive/root:{filePath}:/content";

            //Download the file from OneDrive
            var response = await httpClient.GetAsync(downloadUrl);
            if (response.IsSuccessStatusCode)
            {
                var stream = new MemoryStream();
                await response.Content.CopyToAsync(stream);

                // Reset the stream position to the beginning
                stream.Position = 0;

                Console.WriteLine("File downloaded successfully.");
                return stream;
            }
            else
            {
                Console.WriteLine($"Failed to download file. Status code: {response.StatusCode}");
                string responseBody = await response.Content.ReadAsStringAsync();
                Console.WriteLine($"Error details: {responseBody}");
                return null;
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
