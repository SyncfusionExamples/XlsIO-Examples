using System.IO;
using Syncfusion.XlsIO;
using System;

namespace Reply_Comment
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

                //Access the collection of threaded comments in the worksheet
                IThreadedComments threadedComments = worksheet.ThreadedComments;

                //Add Reply to the Threaded Comment
                threadedComments[0].AddReply("The unit cost of desk is higher compared to other items in the west region. As a result, the total amount is elevated.", "User2", DateTime.Now);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ReplyComment.xlsx"));
                #endregion
            }
        }
    }
}




