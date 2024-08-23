using System;
using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Excel_To_Thumbnail_Image
{
    internal class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet sheet = workbook.Worksheets[0];

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                //Convert to image
                MemoryStream outputStream = new MemoryStream();
                sheet.ConvertToImage(sheet.UsedRange, outputStream);

                //Resize image to thumbnail size
                System.Drawing.Image image = System.Drawing.Image.FromStream(outputStream);
                System.Drawing.Image thumbnail = image.GetThumbnailImage(100, 100, () => false, IntPtr.Zero);

                //Save image
                thumbnail.Save("Image.png", System.Drawing.Imaging.ImageFormat.Png);

            }
        }
    }
}

