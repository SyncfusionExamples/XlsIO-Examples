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
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding comments to a cell
                worksheet.Range["A1"].AddComment().Text = "Comments";

                //Adding comments with author to a cell
                worksheet.Range["A3"].AddComment().Text = worksheet.Range["A3"].Comment.Author;

                //Add Rich Text Comments
                IRange range = worksheet.Range["A6"];
                range.AddComment().RichText.Text = "RichText";
                IRichTextString richText = range.Comment.RichText;

                //Formatting first 4 characters
                IFont redFont = workbook.CreateFont();
                redFont.Bold = true;
                redFont.Color = ExcelKnownColors.Red;
                richText.SetFont(0, 3, redFont);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Comment.xlsx"));
                #endregion
            }
        }
    }
}




