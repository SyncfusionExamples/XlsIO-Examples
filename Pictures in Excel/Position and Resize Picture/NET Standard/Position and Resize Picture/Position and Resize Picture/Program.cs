using System.IO;
using Syncfusion.XlsIO;

namespace Position_and_Resize_Picture
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

                //Adding a picture
                FileStream imageStream = new FileStream("../../../Data/Image.png", FileMode.Open, FileAccess.Read);
                IPictureShape shape = worksheet.Pictures.AddPicture(1, 1, imageStream);

                //Positioning a Picture
                shape.Top = 100;
                shape.Left = 100;

                //Re-sizing a Picture
                shape.Height = 100;
                shape.Width = 100;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("ResizePicture.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                imageStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("ResizePicture.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
