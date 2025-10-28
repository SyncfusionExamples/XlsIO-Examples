using System.IO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO;

namespace Picture_Fill
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
                IChart chart = worksheet.Charts[0];

                //Get data series
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];

                //Getting an image from the stream
                FileStream imageStream1 = new FileStream(Path.GetFullPath(@"Data/Image1.jpg"), FileMode.Open, FileAccess.Read);
                Image image1 = Image.FromStream(imageStream1);
                FileStream imageStream2 = new FileStream(Path.GetFullPath(@"Data/Image2.jpg"), FileMode.Open, FileAccess.Read);
                Image image2 = Image.FromStream(imageStream2);

                //Set picture fill to chart area
                chart.ChartArea.Fill.UserPicture(image1, "Image");

                //Set picture fill to plot area
                chart.PlotArea.Fill.UserPicture(image1, "Image");

                //Set picture fill to series
                serie1.SerieFormat.Fill.UserPicture(image2, "Image");
                serie2.SerieFormat.Fill.UserPicture(image2, "Image");

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
       
                //Dispose streams
                imageStream1.Dispose();
                imageStream2.Dispose();
            }
        }
    }
}





