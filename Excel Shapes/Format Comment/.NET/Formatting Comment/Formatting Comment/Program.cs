using Syncfusion.XlsIO;

namespace Formatting_Comment
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Adding comment in the worksheet with text
                worksheet.Range["A1"].AddComment();
                ICommentShape comment = worksheet.Comments[0];
                comment.Text = "Comment";

                //Set size for the comment
                comment.Height = 150;
                comment.Width = 100;

                //Set position for the comment
                comment.Left = 200;
                comment.Top = 100;

                //Set alignment for the comment
                comment.HAlignment = ExcelCommentHAlign.Right;
                comment.VAlignment = ExcelCommentVAlign.Bottom;

                //Set fill for the comment
                comment.Fill.TwoColorGradient();
                comment.Fill.GradientStyle = ExcelGradientStyle.Horizontal;
                comment.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                comment.Fill.ForeColorIndex = ExcelKnownColors.Red;
                comment.Fill.BackColorIndex = ExcelKnownColors.White;

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}