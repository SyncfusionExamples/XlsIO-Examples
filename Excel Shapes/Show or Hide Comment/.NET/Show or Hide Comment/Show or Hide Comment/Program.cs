using System.IO;
using Syncfusion.XlsIO;

namespace Show_or_Hide_Comment
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

                //Adding comments in the worksheet
                worksheet.Range["E5"].AddComment();
                worksheet.Range["E15"].AddComment();

                //Adding text in comments
                worksheet.Comments[0].Text = "Comment1";
                worksheet.Comments[1].Text = "Comment2";

                //Show comment
                worksheet.Comments[0].IsVisible = true;
                //Hide comment
                worksheet.Comments[1].IsVisible = false;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ShowOrHideComment.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




