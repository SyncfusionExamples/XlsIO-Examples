using Syncfusion.XlsIO;

namespace Chart_Title
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet sheet = workbook.Worksheets[0];
                IChartShape chart = sheet.Charts[0];

                //Set chart name and title
                chart.Name = "Purchase Details";
                chart.ChartTitle = "Purchase Details";

                //Formatting chart title area
                chart.ChartTitleArea.FontName = "Calibri";
                chart.ChartTitleArea.Bold = true;
                chart.ChartTitleArea.Color = ExcelKnownColors.Black;
                chart.ChartTitleArea.Underline = ExcelUnderline.Single;
                chart.ChartTitleArea.Size = 15;

                //Manually resizing chart title area using Layout.
                chart.ChartTitleArea.Layout.Left = 20;

                //Saving the workbook 
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}




