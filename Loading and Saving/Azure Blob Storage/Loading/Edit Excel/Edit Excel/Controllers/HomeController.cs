using Edit_Excel.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using Syncfusion.XlsIO;
using Azure.Storage.Blobs;
using Azure.Storage.Blobs.Models;
using static System.Net.Mime.MediaTypeNames;

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
            // Your Azure Storage Account connection string
            string connectionString = "Your_connection_string";

            // Name of the Azure Blob Storage container
            string containerName = "Your_container_name";

            // Name of the Excel file you want to load
            string blobName = "Your_blob_name";

            // Download the Excel document from Azure Blob Storage
            BlobServiceClient blobServiceClient = new BlobServiceClient(connectionString);
            BlobContainerClient containerClient = blobServiceClient.GetBlobContainerClient(containerName);
            BlobClient blobClient = containerClient.GetBlobClient(blobName);
            try
            {
                // Download the Excel file
                BlobDownloadInfo download = await blobClient.DownloadAsync();

                // Edit the downloaded Excel
                using (Stream fileStream = new MemoryStream())
                {
                    await download.Content.CopyToAsync(fileStream);
                    fileStream.Position = 0;

                    //Create an instance of ExcelEngine
                    using (ExcelEngine excelEngine = new ExcelEngine())
                    {
                        IApplication application = excelEngine.Excel;
                        application.DefaultVersion = ExcelVersion.Excel2016;

                        //Load the downloaded document
                        IWorkbook workbook = application.Workbooks.Open(fileStream);

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
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return Content("Error occurred while processing the file.");
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
