using System.IO;
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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
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
				workbook.SaveAs(Path.GetFullPath("Output/PivotChart.xlsx"));
                #endregion
            }
        }
    }
}





