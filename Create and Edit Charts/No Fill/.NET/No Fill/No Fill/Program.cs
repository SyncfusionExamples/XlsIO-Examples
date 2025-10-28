using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Charts;
using Syncfusion.XlsIO.Implementation.Shapes;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.Drawing;

namespace No_Fill
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];
                IChart chart = worksheet.Charts[0];

                //Get data series
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];

                //Set no fill to chart area
                IChartFrameFormat chartArea = chart.ChartArea;
                chartArea.Fill.Visible = false;

                //Set no fill to plot area
                IChartFrameFormat plotArea = chart.PlotArea;
                plotArea.Fill.Visible = false;

                //Set no fill to series
                serie1.SerieFormat.Fill.Visible = false;

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}




