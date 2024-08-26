using System.IO;
using Syncfusion.XlsIO;
using System;

namespace Resolve_Comment
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/CommentsTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Access the collection of threaded comments in the worksheet
                IThreadedComments threadedComments = worksheet.ThreadedComments;

                //Mark as Resolved
                threadedComments[0].IsResolved = true;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/ResolveComment.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}




