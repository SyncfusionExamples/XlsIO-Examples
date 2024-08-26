using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using Syncfusion.XlsIO.Implementation.Charts;

namespace Data_Labels
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
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];
                IChartShape chart = worksheet.Charts[0];

                //Add the datalabel
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.IsValue = true;

                //Add the datalabel from the range of cells
                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.ValueFromCellsRange = worksheet["I1:I5"];
                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.IsValueFromCells = true;

                //Set the color
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Color = ExcelKnownColors.Blue;
                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.Color = ExcelKnownColors.Black;

                //Set the font
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 10;
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.FontName = "calibri";
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Bold = true;

                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.Size = 10;
                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.FontName = "calibri";
                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.Bold = true;

                //Set the position
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;
                chart.Series[1].DataPoints.DefaultDataPoint.DataLabels.Position = ExcelDataLabelPosition.Outside;

                //Set the number format
                IChartDataLabels dataLabel = chart.Series[0].DataPoints.DefaultDataPoint.DataLabels;
                (dataLabel as ChartDataLabelsImpl).NumberFormat = "#,##0.00";

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}




