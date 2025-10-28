using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Charts;
using Syncfusion.XlsIO.Implementation.Shapes;
using Syncfusion.Drawing;
using Syncfusion.XlsIO.Implementation;

namespace Gradient_Fill
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];
                IChart chart = worksheet.Charts[0];

                //Get data serie
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];

                //Set gradient fill to chart area
                IChartFrameFormat chartArea = chart.ChartArea;
                chartArea.Fill.FillType = ExcelFillType.Gradient;                
                chartArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chartArea.Fill.ForeColor = Color.White;

                //Set gradient fill to plot area
                IChartFrameFormat plotArea = chart.PlotArea;
                plotArea.Fill.FillType = ExcelFillType.Gradient;
                plotArea.Fill.BackColor = Color.FromArgb(205, 217, 234);
                plotArea.Fill.ForeColor = Color.White;

                //Set Gradient fill to series
                ChartFillImpl chartFillImpl1 = serie1.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl1.FillType = ExcelFillType.Gradient;
                chartFillImpl1.GradientColorType = ExcelGradientColor.MultiColor;
                serie1.SerieFormat.Fill.GradientStyle = ExcelGradientStyle.Horizontal;
                GradientStopImpl gradientStopImpl1 = new GradientStopImpl(new ColorObject(Color.FromArgb(0, 176, 240)), 50000, 100000);
                GradientStopImpl gradientStopImpl2 = new GradientStopImpl(new ColorObject(Color.FromArgb(0, 112, 192)), 70000, 100000);
                chartFillImpl1.GradientStops.GradientType = GradientType.Liniar;
                chartFillImpl1.GradientStops.Add(gradientStopImpl1);
                chartFillImpl1.GradientStops.Add(gradientStopImpl2);

                ChartFillImpl chartFillImpl2 = serie2.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl2.FillType = ExcelFillType.Gradient;
                chartFillImpl2.GradientColorType = ExcelGradientColor.MultiColor;
                serie2.SerieFormat.Fill.GradientStyle = ExcelGradientStyle.Horizontal;
                GradientStopImpl gradientStopImpl3 = new GradientStopImpl(new ColorObject(Color.FromArgb(244, 177, 131)), 40000, 100000);
                GradientStopImpl gradientStopImpl4 = new GradientStopImpl(new ColorObject(Color.FromArgb(255, 102, 0)), 70000, 100000);
                chartFillImpl2.GradientStops.GradientType = GradientType.Liniar;
                chartFillImpl2.GradientStops.Add(gradientStopImpl3);
                chartFillImpl2.GradientStops.Add(gradientStopImpl4);

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}





