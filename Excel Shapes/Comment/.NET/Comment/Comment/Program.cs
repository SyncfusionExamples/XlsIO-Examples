using System.IO;
using Syncfusion.XlsIO;

namespace Comment
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
                IWorksheet sheet = workbook.Worksheets[0];

                //Adding comments to a cell
                sheet.Range["A1"].AddComment().Text = "Comments";

                //Add Rich Text Comments
                IRange range = sheet.Range["A6"];
                range.AddComment().RichText.Text = "RichText";
                IRichTextString richText = range.Comment.RichText;

                //Formatting first 4 characters
                IFont redFont = workbook.CreateFont();
                redFont.Bold = true;
                redFont.Color = ExcelKnownColors.Red;
                richText.SetFont(0, 3, redFont);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Comment.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Comment.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
