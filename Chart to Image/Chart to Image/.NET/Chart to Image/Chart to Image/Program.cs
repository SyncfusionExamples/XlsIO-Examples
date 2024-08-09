using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;

namespace Chart_to_Image
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                // Initialize XlsIORenderer
                application.XlsIORenderer = new XlsIORenderer();

                //Set converter chart image format to PNG
                application.XlsIORenderer.ChartRenderingOptions.ImageFormat = ExportImageFormat.Png;

                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                IChart chart = worksheet.Charts[0];

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Image.png", FileMode.Create, FileAccess.Write);
                chart.SaveAsImage(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Image.png")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
