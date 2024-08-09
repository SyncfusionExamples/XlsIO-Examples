using System.IO;
using Syncfusion.XlsIO;

namespace Fill_Comment
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

                //Accessing existing comment
                ICommentShape shape = sheet.Range["A1"].Comment;

                //Format the comment
                shape.Fill.TwoColorGradient();
                shape.Fill.GradientStyle = ExcelGradientStyle.Horizontal;
                shape.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                shape.Fill.ForeColorIndex = ExcelKnownColors.Red;
                shape.Fill.BackColorIndex = ExcelKnownColors.White;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("FillComment.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("FillComment.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
