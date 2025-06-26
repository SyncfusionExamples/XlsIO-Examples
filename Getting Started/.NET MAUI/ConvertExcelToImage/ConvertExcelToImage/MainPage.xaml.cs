using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using System.IO;
using System.Reflection;
using ConvertExcelToImage.SaveServices;

namespace ConvertExcelToImage
{
    public partial class MainPage : ContentPage
    {
        int count = 0;

        public MainPage()
        {
            InitializeComponent();
        }

        private void convertExceltoImage_Click(object sender, EventArgs e)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                Syncfusion.XlsIO.IApplication application = excelEngine.Excel;

                Assembly executingAssembly = typeof(App).GetTypeInfo().Assembly;
                using (Stream inputStream = executingAssembly.GetManifestResourceStream("ConvertExcelToImage.InputTemplate.xlsx"))
                {
                    IWorkbook workbook = application.Workbooks.Open(inputStream);
                    IWorksheet worksheet = workbook.Worksheets[0];

                    //Initialize XlsIO renderer.
                    application.XlsIORenderer = new XlsIORenderer();

                    //Create the MemoryStream to save the image.      
                    MemoryStream imageStream = new MemoryStream();
                    
                    //Save the converted image to MemoryStream.
                    worksheet.ConvertToImage(worksheet.UsedRange, imageStream);
                    imageStream.Position = 0;
                    
                    //save and Launch the Image 
                    SaveService saveService = new();
                    saveService.SaveAndView("ExcelToImage.Jpeg", "application/jpeg", imageStream);
                }
            }
        }
    }
}
