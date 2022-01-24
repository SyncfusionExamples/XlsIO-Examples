using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;

namespace Save_HTML
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Initialize excel engine and open workbook
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];
                worksheet.Range["A1:M20"].Text = "Html Document";

                //Create the instant for SaveOptions
                HtmlSaveOptions saveOptions = new HtmlSaveOptions();
                saveOptions.TextMode = HtmlSaveOptions.GetText.DisplayText;

                #region Save as HTML
                //Saving the workbook
                FileStream outputStream = new FileStream("HTMLFile.html", FileMode.Create, FileAccess.Write);
                workbook.SaveAsHtml(outputStream, saveOptions);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("HTMLFile.html")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
