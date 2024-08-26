using System.IO;
using Syncfusion.XlsIO;

namespace Add_SVG_Picture
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

                FileStream svgStream = new FileStream(Path.GetFullPath(@"Data/Image.svg"), FileMode.Open);
                FileStream pngStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open);

                //Add svg image with given svg and png streams
                worksheet.Pictures.AddPicture(1, 1, svgStream, pngStream);

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/SVGImage.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                svgStream.Dispose();
                pngStream.Dispose();
            }
        }
    }
}





