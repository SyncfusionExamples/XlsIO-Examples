﻿using System.IO;
using Syncfusion.XlsIO;

namespace Create_Pivot_Chart
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[1];
                IPivotTable pivotTable = worksheet.PivotTables[0];

                //Adding a chart to workbook
                IChart pivotChart = workbook.Charts.Add();

                //Set PivotTable as PivotSource to the chart
                pivotChart.PivotSource = pivotTable;

                //Set PivotChart type
                pivotChart.PivotChartType = ExcelChartType.Column_Clustered;

                //Set Field Buttons
                pivotChart.ShowAllFieldButtons = false;
                pivotChart.ShowAxisFieldButtons = false;
                pivotChart.ShowLegendFieldButtons = false;
                pivotChart.ShowReportFilterFieldButtons = false;
                pivotChart.ShowValueFieldButtons = false;

                #region Save
                //Saving the workbook
                FileStream outputStream = new FileStream("PivotChart.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);
                #endregion

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();

                System.Diagnostics.Process process = new System.Diagnostics.Process();
                process.StartInfo = new System.Diagnostics.ProcessStartInfo("PivotChart.xlsx")
                {
                    UseShellExecute = true
                };
                process.Start();
            }
        }
    }
}
