using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Syncfusion.XlsIO;
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Pdf;
using System.IO;
using System.Net.Http.Headers;

namespace Create_Excel_to_PDF
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

                //Initialize ExcelToPdfConverter
                ExcelToPdfConverter excelToPdfConverter = new ExcelToPdfConverter(workbook);

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = excelToPdfConverter.Convert();

                //Create the MemoryStream to save the converted PDF.      
                MemoryStream pdfStream = new MemoryStream();

                //Save the converted PDF document to MemoryStream.
                pdfDocument.Save(pdfStream);
                pdfStream.Position = 0;

                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);

                //Set the PDF document saved stream as content of response.
                response.Content = new ByteArrayContent(pdfStream.ToArray());

                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Sample.pdf"
                };

                //Set the content type as PDF document mime type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

                //Return the response with output PDF document stream.
                return response;
            }
        }
    }
}
