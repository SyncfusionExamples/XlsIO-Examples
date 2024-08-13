using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Custom_Chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                //Merge cells
                sheet.Range["A1:D1"].Merge();

                //Set Font style as bold
                sheet.Range["A1"].CellStyle.Font.Bold = true;

                //Insert data for the chart
                sheet.Range["A1"].Text = "Crescent City, CA";
                sheet.Range["B3"].Text = "Precipitation,in.";
                sheet.Range["C3"].Text = "Temperature,deg.F";
                sheet.Range["A4"].Text = "Jan";
                sheet.Range["A5"].Text = "Feb";
                sheet.Range["A6"].Text = "March";
                sheet.Range["B4"].Number = 10.9;
                sheet.Range["B5"].Number = 8.9;
                sheet.Range["B6"].Number = 8.6;
                sheet.Range["C4"].Number = 47.5;
                sheet.Range["C5"].Number = 48.7;
                sheet.Range["C6"].Number = 48.9;

                //Adjust column width in used range
                sheet.UsedRange.AutofitColumns();

                //Add a new chart with data range
                IChartShape chart = sheet.Charts.Add();
                chart.DataRange = sheet.Range["A3:C6"];

                //Set chart name and chart title
                chart.Name = "CrescentCity,CA";
                chart.ChartTitle = "Crescent City, CA";
                chart.IsSeriesInRows = false;

                //Set primary value axis properties
                chart.PrimaryValueAxis.Title = "Precipitation,in.";
                chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
                chart.PrimaryValueAxis.MaximumValue = 14.0;
                chart.PrimaryValueAxis.NumberFormat = "0.0";

                //Format first serie fill properties
                IChartSerie serieOne = chart.Series[0];
                serieOne.Name = "Precipitation,in.";
                serieOne.SerieFormat.Fill.FillType = ExcelFillType.Gradient;
                serieOne.SerieFormat.Fill.TwoColorGradient(ExcelGradientStyle.Vertical, ExcelGradientVariants.ShadingVariants_2);
                serieOne.SerieFormat.Fill.GradientColorType = ExcelGradientColor.TwoColor;
                serieOne.SerieFormat.Fill.ForeColor = Color.Plum;

                //Format second serie properties
                IChartSerie serieTwo = chart.Series[1];
                serieTwo.SerieType = ExcelChartType.Line_Markers;
                serieTwo.Name = "Temperature,deg.F";

                //Format marker properties
                serieTwo.SerieFormat.MarkerStyle = ExcelChartMarkerType.Diamond;
                serieTwo.SerieFormat.MarkerSize = 8;
                serieTwo.SerieFormat.MarkerBackgroundColor = Color.DarkGreen;
                serieTwo.SerieFormat.MarkerForegroundColor = Color.DarkGreen;
                serieTwo.SerieFormat.LineProperties.LineColor = Color.DarkGreen;

                //Use Secondary Axis
                serieTwo.UsePrimaryAxis = false;

                //MaxCross for secondary axes
                chart.SecondaryCategoryAxis.IsMaxCross = true;
                chart.SecondaryValueAxis.IsMaxCross = true;

                //Set title for secondary value axis
                chart.SecondaryValueAxis.Title = "Temperature,deg.F";
                chart.SecondaryValueAxis.TitleArea.TextRotationAngle = 90;

                //Set secondary category axis properties
                chart.SecondaryCategoryAxis.Border.LineColor = Color.Transparent;
                chart.SecondaryCategoryAxis.MajorTickMark = ExcelTickMark.TickMark_None;
                chart.SecondaryCategoryAxis.TickLabelPosition = ExcelTickLabelPosition.TickLabelPosition_None;

                //Set legend properties
                chart.Legend.Position = ExcelLegendPosition.Bottom;
                chart.Legend.IsVerticalLegend = false;

                //Set Legend
                chart.HasLegend = true;
                chart.Legend.Position = ExcelLegendPosition.Bottom;

                //Positioning the chart in the worksheet
                chart.TopRow = 8;
                chart.LeftColumn = 1;
                chart.BottomRow = 23;
                chart.RightColumn = 8;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("Chart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("Chart.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
