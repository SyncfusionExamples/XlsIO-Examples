using System.IO;
using Syncfusion.XlsIO;
using System;

namespace Add_Comment
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/CommentsTemplate.xlsx"), ExcelOpenType.Automatic);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Add Threaded Comment
                IThreadedComment threadedComment = worksheet.Range["H16"].AddThreadedComment("What is the reason for the higher total amount of \"desk\"  in the west region?", "User1", DateTime.Now);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/AddComment.xlsx"));
                #endregion            
            }
        }
    }
}