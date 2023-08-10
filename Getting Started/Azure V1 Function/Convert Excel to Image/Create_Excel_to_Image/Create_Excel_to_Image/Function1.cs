using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Syncfusion.XlsIO;

namespace Create_Excel_to_Image
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Gets the input Excel document as stream from request.
                Stream excelStream = req.Content.ReadAsStreamAsync().Result;

                //Load the stream into IWorkbook.
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Create the MemoryStream to save the image.      
                MemoryStream imageStream = new MemoryStream();

                //Save the image to MemoryStream.
                worksheet.ConvertToImage(worksheet.UsedRange.Row, worksheet.UsedRange.Column, worksheet.UsedRange.LastRow, worksheet.UsedRange.LastColumn, ImageType.Bitmap, imageStream);
                imageStream.Position = 0;

                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);

                //Set the image saved stream as content of response.
                response.Content = new ByteArrayContent(imageStream.ToArray());

                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Sample.jpeg"
                };

                //Set the content type as image mime type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/jpeg");

                //Return the response with output image stream.
                return response;
            }
        }
    }
}
