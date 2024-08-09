using System.IO;
using Syncfusion.XlsIO;

namespace Add_Picture
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

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("AddPicture.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                imageStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("AddPicture.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
