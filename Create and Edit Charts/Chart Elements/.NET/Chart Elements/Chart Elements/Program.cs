using System.IO;
using Syncfusion.XlsIO;

namespace Chart_Elements
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Positioning chart elements using layout
                //Access the first sheet in the workbook
                IWorksheet sheet = workbook.Worksheets[0];

                IChartShape chart = sheet.Charts[0];

                //Positioning chart in a worksheet
                chart.TopRow = 5;
                chart.LeftColumn = 5;
                chart.RightColumn = 10;
                chart.BottomRow = 10;

                //Manually positioning chart plot area
                chart.PlotArea.Layout.LayoutTarget = LayoutTargets.inner;
                chart.PlotArea.Layout.LeftMode = LayoutModes.edge;
                chart.PlotArea.Layout.TopMode = LayoutModes.edge;

                //Manually positioning chart legend area
                chart.Legend.Layout.LeftMode = LayoutModes.edge;
                chart.Legend.Layout.TopMode = LayoutModes.edge;
                IShape chartShape = chart as IShape;

                //Set Height of the chart in pixels
                chartShape.Height = 300;

                //Set Width of the chart
                chartShape.Width = 500;

                //Manually resizing chart plot area
                chart.PlotArea.Layout.Left = 70;
                chart.PlotArea.Layout.Top = 40;
                chart.PlotArea.Layout.Width = 280;
                chart.PlotArea.Layout.Height = 200;

                //Manually resizing chart legend area
                chart.Legend.Layout.Left = 400;
                chart.Legend.Layout.Top = 150;
                chart.Legend.Layout.Width = 150;
                chart.Legend.Layout.Height = 100;

                // Manually resizing chart title area 
                chart.ChartTitleArea.Text = "Sample Chart";
                chart.ChartTitleArea.Layout.Top = 10;
                chart.ChartTitleArea.Layout.Left = 150;

                // Manually resizing axis title area 
                chart.PrimaryValueAxis.TitleArea.Layout.Left = 15;
                chart.PrimaryValueAxis.TitleArea.Layout.Top = 20;
                chart.PrimaryCategoryAxis.TitleArea.Layout.Left = 25;
                chart.PrimaryCategoryAxis.TitleArea.Layout.Top = 20;

                // Manually resizing data label area 
                chart.Series[0].DataPoints[0].DataLabels.Layout.Left = 0.09;
                chart.Series[0].DataPoints[0].DataLabels.Layout.Top = 0.01;

                //Applying transparency to chart area
                chart.ChartArea.Fill.Transparency = 0.5;

                //Positioning chart elements using manual layout
                //Access the second sheet in the workbook
                IWorksheet sheet1 = workbook.Worksheets[1];

                IChartShape chart1 = sheet1.Charts[0];

                //Positioning chart in a worksheet
                chart.TopRow = 5;
                chart.LeftColumn = 5;
                chart.RightColumn = 10;
                chart.BottomRow = 10;

                //Manually positioning chart plot area
                chart.PlotArea.Layout.ManualLayout.LayoutTarget = LayoutTargets.inner;
                chart.PlotArea.Layout.ManualLayout.LeftMode = LayoutModes.edge;
                chart.PlotArea.Layout.ManualLayout.TopMode = LayoutModes.edge;

                //Manually positioning chart legend area
                chart.Legend.Layout.ManualLayout.LeftMode = LayoutModes.edge;
                chart.Legend.Layout.ManualLayout.TopMode = LayoutModes.edge;
                IShape chartShape1 = chart1 as IShape;

                //Set Height of the chart in pixels
                chartShape.Height = 300;

                //Set Width of the chart
                chartShape.Width = 500;

                //Manually resizing chart plot area
                chart1.PlotArea.Layout.ManualLayout.Height = 0.80;
                chart1.PlotArea.Layout.ManualLayout.Width = 0.65;
                chart1.PlotArea.Layout.ManualLayout.Top = 0.03;
                chart1.PlotArea.Layout.ManualLayout.Left = -0.1;

                //Manually resizing chart legend area
                chart1.Legend.Layout.ManualLayout.Height = 0.09;
                chart1.Legend.Layout.ManualLayout.Width = 0.30;
                chart1.Legend.Layout.ManualLayout.Top = 0.36;
                chart1.Legend.Layout.ManualLayout.Left = 0.68;

                //Manually resizing chart title area
                chart1.ChartTitleArea.Text = "Sample Chart";
                chart1.ChartTitleArea.Layout.ManualLayout.Top = 0.005;
                chart1.ChartTitleArea.Layout.ManualLayout.Left = 0.26;

                //Manually resizing axis title area
                chart1.PrimaryValueAxis.TitleArea.Layout.ManualLayout.Left = 0.04;
                chart1.PrimaryValueAxis.TitleArea.Layout.ManualLayout.Top = 0.34;
                chart1.PrimaryCategoryAxis.TitleArea.Layout.ManualLayout.Left = 0.38;
                chart1.PrimaryCategoryAxis.TitleArea.Layout.ManualLayout.Top = 0.95;

                //Manually resizing data label area
                chart1.Series[0].DataPoints[0].DataLabels.Layout.ManualLayout.Left = 0.09;
                chart1.Series[0].DataPoints[0].DataLabels.Layout.ManualLayout.Top = 0.01;

                //Applying transparency to chart area
                chart1.ChartArea.Fill.Transparency = 0.5;
                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Chart.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





