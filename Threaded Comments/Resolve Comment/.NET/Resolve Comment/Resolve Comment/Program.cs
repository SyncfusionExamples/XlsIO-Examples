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

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/CommentsTemplate.xlsx"), ExcelOpenType.Automatic);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Access the collection of threaded comments in the worksheet
                IThreadedComments threadedComments = worksheet.ThreadedComments;

                //Mark as Resolved
                threadedComments[0].IsResolved = true;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/ResolveComment.xlsx"));
                #endregion
            }
        }
    }
}




