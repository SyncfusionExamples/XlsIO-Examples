
using Syncfusion.XlsIO;
using Syncfusion.Drawing;
using System.IO;

class SalesTrendWithSparklines
{
    static void Main()
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Load the Excel file with sales data
            FileStream inputStream = new FileStream("../../../Data/Sales Data.xlsx", FileMode.Open, FileAccess.Read);

            IWorkbook workbook = application.Workbooks.Open(inputStream);
            IWorksheet sheet = workbook.Worksheets[0];

            sheet["G1"].Text = "Sales Trend";
            sheet["G1"].CellStyle.Font.Bold = true;

            // Create a Sparkline Group
            ISparklineGroup sparklineGroup = sheet.SparklineGroups.Add();

            // Add a Line Sparkline to the group
            ISparklines sparklines = sparklineGroup.Add();
            
            // Define the data range and reference range for the sparkline
            IRange dataRange = sheet.Range["B2:F26"];
            IRange referenceRange = sheet.Range["G2:G26"];
            sparklines.Add(dataRange, referenceRange);

            // Set the sparkline type to Line
            sparklineGroup.SparklineType = SparklineType.Line;

            // Customizing Line Sparkline
            sparklineGroup.LineWeight = 1;
            sparklineGroup.ShowMarkers = true;

            //Set sparkline line color
            sparklineGroup.SparklineColor = Color.Orange;

            // Set the high and low point colors
            sparklineGroup.HighPointColor = Color.Red;
            sparklineGroup.LowPointColor = Color.Green;

            //Customizing markers
            sparklineGroup.MarkersColor = Color.Blue;

            // Save the Excel file
            using (FileStream stream = new FileStream("SalesDataTrend.xlsx", FileMode.Create, FileAccess.Write))
            {
                workbook.SaveAs(stream);
            }
        }
    }
}
