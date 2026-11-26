using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace ImproveExcelToImageQuality
{
    public static class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(@"../../../Data/InputTemplate.xlsx");
                IWorksheet worksheet = workbook.Worksheets[0];

                // Initialize XlsIO renderer.
                application.XlsIORenderer = new XlsIORenderer();

                // Improve quality of the image by setting ScalingMode as Best and ImageFormat as Png which is by default
                ExportImageOptions exportImageOptions = new ExportImageOptions();
                exportImageOptions.ScalingMode = ScalingMode.Best;
                exportImageOptions.ImageFormat = ExportImageFormat.Png;
                
                // Saving the excel as image
                FileStream outputStream = new FileStream(@"../../../Output/Image.png", FileMode.Create, FileAccess.Write);
                worksheet.ConvertToImage(worksheet.UsedRange, outputStream);

                outputStream.Dispose();
                workbook.Close();
            }
        }
    }
}