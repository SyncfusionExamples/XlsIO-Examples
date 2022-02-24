using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Comments
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Remove all the comments in worksheet
                worksheet.Comments.Clear();

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("RemoveComments.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("RemoveComments.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
