using Syncfusion.Drawing;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System.IO;
using static System.Net.Mime.MediaTypeNames;

namespace ConvertExceltoImage.Components.Data
{
    public class ExcelService
    {
        public MemoryStream ExceltoImage()
        {          
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream excelStream = new FileStream("InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);
                IWorksheet worksheet = workbook.Worksheets[0];
                //Initialize XlsIO renderer.
                application.XlsIORenderer = new XlsIORenderer();
                //Create the MemoryStream to save the image.
                MemoryStream imageStream = new MemoryStream();

                //Save the converted image to MemoryStream.
                worksheet.ConvertToImage(worksheet.UsedRange, imageStream);
                imageStream.Position = 0;

                //Download image in the browser.
                return imageStream;
            }            
        }
    }
}

