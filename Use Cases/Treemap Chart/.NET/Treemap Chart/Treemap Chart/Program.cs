using Syncfusion.XlsIO;

namespace Treemap_Chart
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

                //Enter sample data
                worksheet.Range["A1"].Text = "Category";
                worksheet.Range["B1"].Text = "SubCategory";
                worksheet.Range["C1"].Text = "Value";

                worksheet.Range["A2"].Text = "Fruit";
                worksheet.Range["B2"].Text = "Apple";
                worksheet.Range["C2"].Number = 50;

                worksheet.Range["A3"].Text = "Fruit";
                worksheet.Range["B3"].Text = "Banana";
                worksheet.Range["C3"].Number = 30;

                worksheet.Range["A4"].Text = "Vegetable";
                worksheet.Range["B4"].Text = "Carrot";
                worksheet.Range["C4"].Number = 40;

                worksheet.Range["A5"].Text = "Vegetable";
                worksheet.Range["B5"].Text = "Broccoli";
                worksheet.Range["C5"].Number = 25;

                //Add chart to worksheet
                IChartShape chart = worksheet.Charts.Add();

                //Set chart type to Treemap
                chart.ChartType = ExcelChartType.TreeMap;

                //Set chart data range
                chart.DataRange = worksheet.Range["A1:C5"];
                chart.IsSeriesInRows = false;

                //Set chart title
                chart.ChartTitle = "Treemap Chart";

                //Positioning the chart in the worksheet
                chart.TopRow = 8;
                chart.LeftColumn = 1;
                chart.BottomRow = 23;
                chart.RightColumn = 8;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}
