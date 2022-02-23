using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Linked_OLE
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Create image stream
                FileStream imageStream = new FileStream("../../../Data/Image.png", FileMode.Open);
                //Get image from stream
                Image image = Image.FromStream(imageStream);

                //Add ole object
                IOleObject oleObject = worksheet.OleObjects.AddLink("../../../Data/Document.docx", image);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("LinkedOLE.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                imageStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("LinkedOLE.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
