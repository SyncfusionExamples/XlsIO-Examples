using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Embedded_OLE
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

                //Create file stream and image stream
                FileStream inputStream = new FileStream("../../../Data/Presentation Document.pptx", FileMode.Open);
                FileStream imageStream = new FileStream("../../../Data/Image.png", FileMode.Open);

                //Get image from stream
                Image image = Image.FromStream(imageStream);

                //Add ole object
                IOleObject oleObject = worksheet.OleObjects.Add(inputStream, image, OleObjectType.PowerPointPresentation);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("EmbeddedOLE.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
                imageStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("EmbeddedOLE.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
