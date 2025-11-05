using Syncfusion.Licensing;
using Syncfusion.XlsIO;

namespace Delete_Comment
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/CommentsTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the collection of threaded comments in the worksheet
                IThreadedComments threadedComments = worksheet.ThreadedComments;

                //Delete the threaded comment
                threadedComments[0].Delete();

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/DeleteComment.xlsx"));
            }
        }
    }
}