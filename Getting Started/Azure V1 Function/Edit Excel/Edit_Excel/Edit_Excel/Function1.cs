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

namespace Edit_Excel
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Gets the input Excel document as stream from request.
                Stream inputStream = req.Content.ReadAsStreamAsync().Result;

                //Load the stream into IWorkbook.
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Get the first worksheet in the workbook into IWorksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Assign some text in a cell
                worksheet.Range["A3"].Text = "Hello World";

                //Access a cell value from Excel
                var value = worksheet.Range["A1"].Value;

                //Create the MemoryStream to save the Excel.      
                MemoryStream excelStream = new MemoryStream();

                //Save the Excel document to MemoryStream.
                workbook.SaveAs(excelStream);
                excelStream.Position = 0;

                //Create the response to return.
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);

                //Set the Excel document saved stream as content of response.
                response.Content = new ByteArrayContent(excelStream.ToArray());

                //Set the contentDisposition as attachment.
                response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
                {
                    FileName = "Output.xlsx"
                };

                //Set the content type as Excel document mime type.
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/excel");

                //Return the response with output PDF document stream.
                return response;
            }
        }
    }
}
