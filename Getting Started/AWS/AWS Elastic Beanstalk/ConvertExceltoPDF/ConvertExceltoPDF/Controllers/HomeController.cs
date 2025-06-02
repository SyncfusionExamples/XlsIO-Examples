using ConvertExceltoPDF.Models;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.Drawing;
using Syncfusion.Pdf;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIORenderer;
using System.Diagnostics;
using System.IO;

namespace ConvertExceltoPDF.Controllers
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

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult ConvertExceltoPDF()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Initializes the SubstituteFont event to perform font substitution during Excel-to-PDF conversion
                application.SubstituteFont += new SubstituteFontEventHandler(SubstituteFont);

                FileStream excelStream = new FileStream(@"Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(excelStream);

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Convert Excel document into PDF document 
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                //Create the MemoryStream to save the converted PDF.      
                MemoryStream pdfStream = new MemoryStream();

                //Save the converted PDF document to MemoryStream.
                pdfDocument.Save(pdfStream);
                pdfStream.Position = 0;
                //Download PDF document in the browser
                return File(pdfStream, "application/pdf", "Sample.pdf");
            }
        }

        private static void SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            string filePath = string.Empty;
            FileStream fileStream = null;

            if (args.OriginalFontName == "Calibri")
            {
                filePath = Path.GetFullPath(@"Data/calibri.ttf");
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                args.AlternateFontStream = fileStream;
            }
        }
    }
}
