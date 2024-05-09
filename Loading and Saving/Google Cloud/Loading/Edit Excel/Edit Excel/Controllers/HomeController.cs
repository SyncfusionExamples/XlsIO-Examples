using Edit_Excel.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.XlsIO;
using Google.Cloud.Storage.V1;
using Google.Apis.Auth.OAuth2;

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
            //Your bucket name
            string bucketName = "Your_bucket_name";

            //Your service account key path
            string keyPath = "Your_service_account_key_path";

            //Name of the file to download from the Google Cloud Storage
            string fileName = "Your_file_name";

            //Create Google Credential from the service account key file
            GoogleCredential credential = GoogleCredential.FromFile(keyPath);

            //Instantiates a storage client to interact with Google Cloud Storage
            StorageClient storageClient = StorageClient.Create(credential);

            //Download a file from Google Cloud Storage
            using (MemoryStream memoryStream = new MemoryStream())
            {
                await storageClient.DownloadObjectAsync(bucketName, fileName, memoryStream);
                memoryStream.Position = 0;

                //Edit the downloaded Excel file
                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //Loads the downloaded document
                    IWorkbook workbook = application.Workbooks.Open(memoryStream);

                    IWorksheet worksheet = workbook.Worksheets[0];
                    worksheet.Range["A3"].Text = "Hello world";

                    //Saving the Excel to the MemoryStream 
                    MemoryStream outputStream = new MemoryStream();
                    workbook.SaveAs(outputStream);

                    //Set the position as '0'.
                    outputStream.Position = 0;

                    //Download the Excel file in the browser
                    FileStreamResult fileStreamResult = new FileStreamResult(outputStream, "application/excel");
                    fileStreamResult.FileDownloadName = "EditExcel.xlsx";
                    return fileStreamResult;
                }
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
