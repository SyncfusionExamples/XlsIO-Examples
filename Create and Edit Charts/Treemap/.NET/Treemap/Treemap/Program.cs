﻿using System.IO;
using Syncfusion.XlsIO;

namespace Treemap
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
                IWorksheet sheet = workbook.Worksheets[0];

                //Create a chart
                IChartShape chart = sheet.Charts.Add();

                //Set chart type as TreeMap
                chart.ChartType = ExcelChartType.TreeMap;

                //Set data range in the worksheet
                chart.DataRange = sheet["A2:C11"];

                //Set the chart title
                chart.ChartTitle = "Area by countries";

                //Set the Treemap label option
                chart.Series[0].SerieFormat.TreeMapLabelOption = ExcelTreeMapLabelOption.Banner;

                //Formatting data labels      
                chart.Series[0].DataPoints.DefaultDataPoint.DataLabels.Size = 8;

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Treemap.xlsx"));
                #endregion
            }
        }
    }
}





