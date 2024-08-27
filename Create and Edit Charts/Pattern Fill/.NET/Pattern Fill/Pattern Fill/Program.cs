using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Charts;
using Syncfusion.Drawing;
using Image = Syncfusion.Drawing.Image;

namespace Pattern_Fill
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];
                IChart chart = worksheet.Charts[0];

                //Get data series
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];

                //Set pattern fill to chart area
                IChartFrameFormat chartArea = chart.ChartArea;
                chartArea.Fill.FillType = ExcelFillType.Pattern;
                chartArea.Fill.BackColor = Color.Pink;
                chartArea.Fill.ForeColor = Color.White;
                chartArea.Fill.Pattern = ExcelGradientPattern.Pat_90_Percent;

                //Set pattern fill to plot area
                IChartFrameFormat plotArea = chart.PlotArea;
                plotArea.Fill.FillType = ExcelFillType.Pattern;
                plotArea.Fill.BackColor = Color.Pink;
                plotArea.Fill.ForeColor = Color.White;
                plotArea.Fill.Pattern = ExcelGradientPattern.Pat_90_Percent;

                //Set pattern fill to series
                ChartFillImpl chartFillImpl1 = serie1.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl1.FillType = ExcelFillType.Pattern;
                chartFillImpl1.BackColor = Color.Pink;
                chartFillImpl1.ForeColor = Color.White;
                chartFillImpl1.Pattern = ExcelGradientPattern.Pat_5_Percent;

                ChartFillImpl chartFillImpl2 = serie2.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl2.FillType = ExcelFillType.Pattern;
                chartFillImpl2.BackColor = Color.Gray;
                chartFillImpl2.ForeColor = Color.White;
                chartFillImpl2.Pattern = ExcelGradientPattern.Pat_5_Percent;

                //Saving the workbook as stream
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




