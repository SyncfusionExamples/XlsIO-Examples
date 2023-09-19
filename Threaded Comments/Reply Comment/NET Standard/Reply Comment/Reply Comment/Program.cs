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

                FileStream inputStream = new FileStream("../../../Data/CommentsTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);

                IWorksheet worksheet = workbook.Worksheets[0];

                //Add Threaded Comment
                IThreadedComment threadedComment = worksheet.Range["H16"].AddThreadedComment("What is the reason for the higher total amount of \"desk\"  in the west region?", "User1", DateTime.Now);

                //Add Reply to the Threaded Comment
                threadedComment.AddReply("The unit cost of desk is higher compared to other items in the west region. As a result, the total amount is elevated.", "User2", DateTime.Now);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ReplyComment.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ReplyComment.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}