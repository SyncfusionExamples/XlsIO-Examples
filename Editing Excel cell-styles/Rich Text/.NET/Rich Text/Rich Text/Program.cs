using System.IO;
using Syncfusion.XlsIO;

namespace Rich_Text
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Add Text
                IRange range = worksheet.Range["A1"];
                range.Text = "RichText";
                IRichTextString richText = range.RichText;

                //Formatting first 4 characters.
                IFont redFont = workbook.CreateFont();
                redFont.Bold = true;
                redFont.Italic = true;
                redFont.RGBColor = Syncfusion.Drawing.Color.Red;
                richText.SetFont(0, 3, redFont);

                //Formatting last 4 characters.
                IFont blueFont = workbook.CreateFont();
                blueFont.Bold = true;
                blueFont.Italic = true;
                blueFont.RGBColor = Syncfusion.Drawing.Color.Blue;
                richText.SetFont(4, 7, blueFont);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("RichText.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("RichText.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
