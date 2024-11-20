using Syncfusion.Licensing;
using Syncfusion.XlsIO;

namespace Import_Collection
{
    class Program
    {
        public static void Main(string[] args)
        {
            // Initialize Excel engine and application.
            using ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            // Open an existing workbook.
            FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
            IWorkbook workbook = application.Workbooks.Open(inputStream);

            // Create Template Marker Processor.
            ITemplateMarkersProcessor marker = workbook.CreateTemplateMarkersProcessor();

            // Get the data into collection object.
            IList<Report> reports = GetSalesReports();

            // Add collections to the marker variables where the name should match with input template.
            marker.AddVariable("Reports", reports);

            //Applying Markers
            marker.ApplyMarkers();

            // Saving the workbook.
            FileStream outputStream = new FileStream(Path.GetFullPath("Output/ImportCollection.xlsx"), FileMode.Create, FileAccess.Write);
            workbook.SaveAs(outputStream);

            //Dispose streams
            inputStream.Dispose();
            outputStream.Dispose();
        }
        // Gets a list of sales reports.
        private static List<Report> GetSalesReports()
        {
            List<Report> reports = new List<Report>();
            reports.Add(new Report("Andy Bernard", "45000", "58000", 29, "Data/Andy.png"));
            reports.Add(new Report("Jim Halpert", "34000", "65000", 91, "Data/Jim.png"));
            reports.Add(new Report("Karen Fillippelli", "75000", "64000", -14, "Data/Karen.png"));
            reports.Add(new Report("Phyllis Lapin", "56500", "33600", -40, "Data/Phyllis.png"));
            reports.Add(new Report("Stanley Hudson", "46500", "52000", 12, "Data/Stanley.png"));
            return reports;
        }

        // Sales report.
        public class Report
        {
            public string SalesPerson { get; set; }
            public string SalesJanJun { get; set; }
            public string SalesJulDec { get; set; }
            public int Change { get; set; }
            public byte[] Image { get; set; }

            public Report(string name, string janToJun, string julToDec, int change, string imagePath)
            {
                SalesPerson = name;
                SalesJanJun = janToJun;
                SalesJulDec = julToDec;
                Change = change;
                Image = File.ReadAllBytes(imagePath);
            }
        }
    }
}