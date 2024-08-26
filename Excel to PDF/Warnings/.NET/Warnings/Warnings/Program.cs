using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

namespace Warnings
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                //Open the Excel document to convert.
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Initialize warning class to capture warnings during the conversion.
                Warning warning = new Warning();

                //Initialize XlsIO renderer.
                XlsIORenderer renderer = new XlsIORenderer();

                //Initialize XlsIO renderer settings.
                XlsIORendererSettings settings = new XlsIORendererSettings();

                //Set the warning class that is implemented.
                settings.Warning = warning;

                //Convert Excel document into PDF document.
                PdfDocument pdfDocument = renderer.ConvertToPDF(workbook, settings);

                //If conversion process canceled null returned.
                if (pdfDocument != null)
                {
                    #region Save
                    //Saving the workbook
                    FileStream outputStream = new FileStream(Path.GetFullPath("Output/ExceltoPDF.pdf"), FileMode.Create, FileAccess.Write);
                    pdfDocument.Save(outputStream);
                    #endregion

                    //Dispose streams
                    outputStream.Dispose();
                }
                inputStream.Dispose();
            }
        }
    }
    public class Warning : IWarning
    {
        public void ShowWarning(WarningInfo warning)
        {
            //Cancel the converion process if the warning type is FillPattern.
            if (warning.Type == WarningType.FillPattern)
                Cancel = true;

            //To view or log the warning, you can make use of warning.Description.
        }
        public bool Cancel { get; set; }
    }
}





