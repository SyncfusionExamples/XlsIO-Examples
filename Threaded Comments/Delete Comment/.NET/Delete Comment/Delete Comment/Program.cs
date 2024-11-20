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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/CommentsTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the collection of threaded comments in the worksheet
                IThreadedComments threadedComments = worksheet.ThreadedComments;

                //Delete the threaded comment
                threadedComments[0].Delete();

                //Saving the workbook as stream
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/DeleteComment.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}