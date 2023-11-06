using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using System.Reflection;

namespace Convert_Excel_to_PDF;

public partial class MainPage : ContentPage
{
	int count = 0;
	public MainPage()
	{
		InitializeComponent();
	}

	private void ConvertExceltoPDF(object sender, EventArgs e)
	{
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            Syncfusion.XlsIO.IApplication application = excelEngine.Excel;

            Assembly executingAssembly = typeof(App).GetTypeInfo().Assembly;
            using (Stream inputStream = executingAssembly.GetManifestResourceStream("Convert_Excel_to_PDF.InputTemplate.xlsx"))
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