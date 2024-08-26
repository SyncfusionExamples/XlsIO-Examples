using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Format_Series
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["A1"].Text = "Items";
                worksheet.Range["A2"].Text = "Beverages";
                worksheet.Range["A3"].Text = "Condiments";
                worksheet.Range["A4"].Text = "Confections";
                worksheet.Range["A5"].Text = "Dairy Products";
                worksheet.Range["A6"].Text = "Grains/Cereals";

                worksheet.Range["B1"].Text = "Amount(in $)";
                worksheet.Range["B2"].Number = 2776;
                worksheet.Range["B3"].Number = 1077;
                worksheet.Range["B4"].Number = 2287;
                worksheet.Range["B5"].Number = 1368;
                worksheet.Range["B6"].Number = 3325;

                worksheet.Range["C1"].Text = "Count";
                worksheet.Range["C2"].Number = 925;
                worksheet.Range["C3"].Number = 378;
                worksheet.Range["C4"].Number = 880;
                worksheet.Range["C5"].Number = 581;
                worksheet.Range["C6"].Number = 189;

                IChartShape chart = worksheet.Charts.Add();

                //Set chart type
                chart.ChartType = ExcelChartType.Column_Clustered;

                //Set chart title
                chart.ChartTitle = "Product Sales";

                //Add first serie
                IChartSerie serie1 = chart.Series.Add("Amount");
                serie1.Values = worksheet.Range["B2:B6"];
                serie1.CategoryLabels = worksheet.Range["A2:A6"];

                //Add second serie
                IChartSerie serie2 = chart.Series.Add("Count");
                serie2.Values = worksheet.Range["C2:C6"];
                serie2.CategoryLabels = worksheet.Range["A2:A6"];

                //Set the series type
                chart.Series[0].SerieType = ExcelChartType.Line_Markers;
                chart.Series[1].SerieType = ExcelChartType.Bar_Clustered;

                //Set the color
                chart.Series[1].SerieFormat.Fill.FillType = ExcelFillType.Gradient;
                chart.Series[1].SerieFormat.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                chart.Series[1].SerieFormat.Fill.BackColor = Color.FromArgb(205, 217, 234);
                chart.Series[1].SerieFormat.Fill.ForeColor = Color.Red;

                //Set the border
                chart.Series[1].SerieFormat.LineProperties.LineColor = Color.Red;
                chart.Series[1].SerieFormat.LineProperties.LinePattern = ExcelChartLinePattern.Dot;
                chart.Series[1].SerieFormat.LineProperties.LineWeight = ExcelChartLineWeight.Narrow;

                //Positioning chart in a worksheet
                chart.TopRow = 9;
                chart.LeftColumn = 1;
                chart.RightColumn = 10;
                chart.BottomRow = 25;

                //Saving the workbook as stream
                FileStream stream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);

                //Dispose streams
                stream.Dispose();
            }
        }
    }
}



