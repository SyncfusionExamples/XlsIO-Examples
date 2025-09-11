using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Chart_to_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            //Initialize ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Initialize application
                IApplication application = excelEngine.Excel;

                //Set the default version as Xlsx
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                //Set converter chart image format to PNG or JPEG
                application.XlsIORenderer.ChartRenderingOptions.ImageFormat = ExportImageFormat.Png;

                //Set the chart image quality to best
                application.XlsIORenderer.ChartRenderingOptions.ScalingMode = ScalingMode.Best;

                //Open existing workbook with chart
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Access the chart from the worksheet
                IChart chart = worksheet.Charts[0];

                #region Save
                //Exporting the chart as image
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Image.png"), FileMode.Create, FileAccess.Write);
                chart.SaveAsImage(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





