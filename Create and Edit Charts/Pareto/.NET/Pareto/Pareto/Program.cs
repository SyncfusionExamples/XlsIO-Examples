using System.IO;
using Syncfusion.XlsIO;

namespace Pareto
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
                IWorkbook workbook = application.Workbooks.Open(inputStream, ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a chart
                IChartShape chart = sheet.Charts.Add();

                //Set chart type as Pareto
                chart.ChartType = ExcelChartType.Pareto;

                //Set data range in the worksheet   
                chart.DataRange = sheet["A2:B8"];

                //Set category values as bin values   
                chart.PrimaryCategoryAxis.IsBinningByCategory = true;

                //Formatting Pareto line      
                chart.Series[0].ParetoLineFormat.LineProperties.ColorIndex = ExcelKnownColors.Bright_green;

                //Gap width settings
                chart.Series[0].SerieFormat.CommonSerieOptions.GapWidth = 6;

                //Set the chart title
                chart.ChartTitle = "Expenses";

                //Hiding the legend
                chart.HasLegend = false;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Pareto.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}





