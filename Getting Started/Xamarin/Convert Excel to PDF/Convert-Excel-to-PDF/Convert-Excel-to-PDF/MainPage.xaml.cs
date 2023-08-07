using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Forms;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using System.IO;
using System.Reflection;

namespace Convert_Excel_to_PDF
{
    public partial class MainPage : ContentPage
    {
        public MainPage()
        {
            InitializeComponent();
        }
        private void OnButtonClicked(object sender, EventArgs e)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                Syncfusion.XlsIO.IApplication application = excelEngine.Excel;

                Assembly executingAssembly = typeof(App).GetTypeInfo().Assembly;
                using(Stream inputStream = executingAssembly.GetManifestResourceStream("Convert-Excel-to-PDF.InputTemplate.xlsx"))
                {
                    // Open the workbook.
                    IWorkbook workbook = application.Workbooks.Open(inputStream);

                    // Instantiate the Excel to PDF renderer.
                    XlsIORenderer renderer = new XlsIORenderer();

                    //Convert Excel document into PDF document 
                    PdfDocument pdfDocument = renderer.ConvertToPDF(workbook);

                    //Create the MemoryStream to save the converted PDF.      
                    MemoryStream pdfStream = new MemoryStream();

                    //Save the converted PDF document to MemoryStream.
                    pdfDocument.Save(pdfStream);
                    pdfStream.Position = 0;

                    //Save the stream as a file in the device and invoke it for viewing.
                    Xamarin.Forms.DependencyService.Get<ISave>().SaveAndView("Sample.pdf", "application/pdf", pdfStream);
                }
            }
        }
    }
}
