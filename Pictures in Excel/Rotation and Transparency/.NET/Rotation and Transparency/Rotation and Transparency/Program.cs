using Syncfusion.XlsIO;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace Rotation_and_Transparency
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing Excel file into IWorkbook
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Sample.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                Image image = Image.FromFile(Path.GetFullPath(@"Data/image.png"));
                Bitmap bitmap = new Bitmap(image);
                bitmap.RotateFlip(RotateFlipType.Rotate90FlipNone);
                bitmap.MakeTransparent(Color.Black);
                bitmap.Save("image_M.png", ImageFormat.Png);

                FileStream imageStream = new FileStream("image_M.png", FileMode.Open, FileAccess.Read);
                worksheet.PageSetup.BackgoundImage = Syncfusion.Drawing.Image.FromStream(imageStream);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}

