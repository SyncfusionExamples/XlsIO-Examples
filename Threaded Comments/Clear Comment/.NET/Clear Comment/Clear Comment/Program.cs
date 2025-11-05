using Syncfusion.XlsIO;

namespace Clear_Comment
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

                //Clear all the threaded comments
                threadedComments.Clear();

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ClearComment.xlsx"));
            }
        }
    }
}
