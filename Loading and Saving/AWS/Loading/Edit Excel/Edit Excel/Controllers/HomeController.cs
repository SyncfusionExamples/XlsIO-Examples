using Amazon.S3.Transfer;
using Amazon.S3;
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

        public IActionResult Index()
        {
            return View();
        }
        public async Task<IActionResult> EditDocument()
        {
            //Your AWS Storage Account bucket name 
            string bucketName = "your-bucket-name";

            //Name of the Excel file you want to load from AWS S3
            string key = "CreateExcel.xlsx";

            // Configure AWS credentials and region
            var region = Amazon.RegionEndpoint.USEast1; 
            var credentials = new Amazon.Runtime.BasicAWSCredentials("your-access-key", "your-secret-key"); 
            var config = new AmazonS3Config
            {
                RegionEndpoint = region
            };

            try
            {
                using (var client = new AmazonS3Client(credentials, config))
                {
                    // Create a MemoryStream to copy the file content
                    using (MemoryStream stream = new MemoryStream())
                    {
                        // Download the file from S3 into the MemoryStream
                        var response = await client.GetObjectAsync(new Amazon.S3.Model.GetObjectRequest
                        {
                            BucketName = bucketName,
                            Key = key
                        });

                        // Copy the response stream to the MemoryStream
                        await response.ResponseStream.CopyToAsync(stream);

                        // Set the position to the beginning of the MemoryStream
                        stream.Position = 0;

                        //Create an instance of ExcelEngine
                        using (ExcelEngine excelEngine = new ExcelEngine())
                        {
                            IApplication application = excelEngine.Excel;
                            application.DefaultVersion = ExcelVersion.Excel2016;

                            //Load the downloaded document
                            IWorkbook workbook = application.Workbooks.Open(stream);

                            //Access the first worksheet
                            IWorksheet worksheet = workbook.Worksheets[0];

                            //Modify the text
                            worksheet.Range["A3"].Text = "Hello world";

                            //Saving the workbook to the MemoryStream 
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
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return Content("Error occurred while processing the file.");
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
