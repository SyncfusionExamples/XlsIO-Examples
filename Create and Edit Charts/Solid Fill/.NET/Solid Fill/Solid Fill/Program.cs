﻿using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Charts;
using Syncfusion.XlsIO.Implementation.Shapes;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.Drawing;

namespace Solid_Fill
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];
                IChart chart = worksheet.Charts[0];
                
                //Get data series
                IChartSerie serie1 = chart.Series[0];
                IChartSerie serie2 = chart.Series[1];

                //Set solid fill to chart area
                IChartFrameFormat chartArea = chart.ChartArea;
                chartArea.Fill.FillType = ExcelFillType.SolidColor;
                chartArea.Fill.ForeColor = Color.FromArgb(208,206,206);

                //Set solid fill to plot area
                IChartFrameFormat plotArea = chart.PlotArea;
                plotArea.Fill.FillType = ExcelFillType.SolidColor;
                plotArea.Fill.ForeColor = Color.FromArgb(208, 206, 206);

                //Set solid fill to series
                ChartFillImpl chartFillImpl1 = serie1.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl1.FillType = ExcelFillType.SolidColor;
                chartFillImpl1.ForeColor = Color.FromArgb(255, 192, 203);

                ChartFillImpl chartFillImpl2 = serie2.SerieFormat.Fill as ChartFillImpl;
                chartFillImpl2.FillType = ExcelFillType.SolidColor;
                chartFillImpl2.ForeColor = Color.FromArgb(143, 170, 220); ;

                //Saving the workbook as streams
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Output.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}