using System;
using System.IO;
using Syncfusion.ExcelToPdfConverter;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using System.Web;

namespace Convert_Excel_to_PDF
{
    public partial class MainPage : System.Web.UI.Page
    {
        protected void OnButtonClicked(object sender, EventArgs e)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing file
                FileStream excelStream = new FileStream(Server.MapPath("~/App_Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                //Initialize ExcelToPdfConverter
                ExcelToPdfConverter converter = new ExcelToPdfConverter(workbook);

                //Initialize PDF document
                PdfDocument pdfDocument = new PdfDocument();

                //Convert Excel document into PDF document
                pdfDocument = converter.Convert();

                //Create the MemoryStream to save the converted PDF.      
                MemoryStream pdfStream = new MemoryStream();

                //Save the converted PDF document to MemoryStream.
                pdfDocument.Save("sample.pdf", HttpContext.Current.Response, HttpReadType.Save);
            }
        }
    }
}