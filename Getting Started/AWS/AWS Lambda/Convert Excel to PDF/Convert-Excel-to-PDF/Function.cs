using Amazon.Lambda.Core;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;
using Syncfusion.XlsIO.Implementation;
// Assembly attribute to enable the Lambda function's JSON input to be converted into a .NET class.
[assembly: LambdaSerializer(typeof(Amazon.Lambda.Serialization.SystemTextJson.DefaultLambdaJsonSerializer))]

namespace Convert_Excel_to_PDF
{
    public class Function
    {

        /// <summary>
        /// A simple function that takes a string and does a ToUpper
        /// </summary>
        /// <param name="input"></param>
        /// <param name="context"></param>
        /// <returns></returns>
        public string FunctionHandler(string input, ILambdaContext context)
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
                return Convert.ToBase64String(pdfStream.ToArray());
            }
        }
        private void SubstituteFont(object sender, SubstituteFontEventArgs args)
        {
            string filePath = string.Empty;
            FileStream fileStream = null;

            if (args.OriginalFontName == "Calibri")
            {
                filePath = Path.GetFullPath(@"Data/calibri.ttf");
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                args.AlternateFontStream = fileStream;
            }
            else if (args.OriginalFontName == "Arial")
            {
                filePath = Path.GetFullPath(@"Data/arial.ttf");
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                args.AlternateFontStream = fileStream;
            }
            else
            {
                filePath = Path.GetFullPath(@"Data/times.ttf");
                fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                args.AlternateFontStream = fileStream;
            }
        }
    }
}

