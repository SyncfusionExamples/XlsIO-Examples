using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System.Reflection;
using Syncfusion.Pdf;

namespace ConvertExceltoPDF
{
    public partial class MainPage : ContentPage
    {

        public MainPage()
        {
            InitializeComponent();
        }

        private void convertExceltoPDF_Click(object sender, EventArgs e)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                Syncfusion.XlsIO.IApplication application = excelEngine.Excel;

                Assembly executingAssembly = typeof(App).GetTypeInfo().Assembly;
                using (Stream inputStream = executingAssembly.GetManifestResourceStream("ConvertExceltoPDF.InputTemplate.xlsx"))
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

                    //save and Launch the PDF document
                    SaveService saveService = new();
                    saveService.SaveAndView("Sample.pdf", "application/pdf", pdfStream);
                }
            }
        }
    }
}
