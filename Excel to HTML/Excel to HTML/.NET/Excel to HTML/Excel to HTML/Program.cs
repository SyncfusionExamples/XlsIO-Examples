using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;

namespace Excel_to_HTML
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
                worksheet.UsedRange.AutofitColumns();

                #region Save as HTML
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.html"), FileMode.Create, FileAccess.Write);
                workbook.SaveAsHtml(outputStream, saveOptions);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




