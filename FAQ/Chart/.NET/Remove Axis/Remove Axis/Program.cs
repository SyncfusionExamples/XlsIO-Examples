using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        // Create an instance of ExcelEngine
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;

            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data\InputTemplate.xlsx"));

            IWorksheet worksheet = workbook.Worksheets[0];

            IChartShape chart = worksheet.Charts[0];

            //Removing legend.
            chart.HasLegend = false;

            //Removing CategoryAxis. 
            chart.PrimaryCategoryAxis.Visible = false;

            //Removing ValueAxix.
            chart.PrimaryValueAxis.Visible = false;
            // Save the workbook to a file
            workbook.SaveAs(Path.GetFullPath(@"Output\Output.xlsx"));
        }
    }
}