using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Hyperlink
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                #region Remove Hyperlink
                //Removing Hyperlink from Range "C7"
                worksheet.Range["C7"].Hyperlinks.RemoveAt(0);
                #endregion

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("RemoveHyperlink.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("RemoveHyperlink.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
