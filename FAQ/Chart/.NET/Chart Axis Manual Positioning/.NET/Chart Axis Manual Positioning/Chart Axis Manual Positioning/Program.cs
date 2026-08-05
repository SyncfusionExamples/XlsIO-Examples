using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IWorkbook workbook = excelEngine.Excel.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
            IWorksheet worksheet = workbook.Worksheets[0];

            // Access the chart in the worksheet
            IChartShape chart = worksheet.Charts[0];

            chart.PrimaryValueAxis.TitleArea.Layout.ManualLayout.Top = 0.15; // Adjust the vertical position of the title


            // Save the modified workbook
            workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
        }
    }
}