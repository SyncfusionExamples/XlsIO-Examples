using Syncfusion.XlsIO;
using System.ComponentModel;

namespace Export_delimited_string_to_string_array 
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an input template
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Export worksheet data into Collection Objects
                List<Report> collectionObjects = worksheet.ExportData<Report>(1, 1, 6, 3);

                //Loop through the list and add the delimitied string to an array
                foreach (var report in collectionObjects)
                {
                    report.SalesPersonArray = report.SalesPerson.Split(',');
                }

                //Dispose streams
                inputStream.Dispose();

            }
        }
        public class Report
        {
            [DisplayNameAttribute("Sales Person Name")]
            public string SalesPerson { get; set; }
            public string SalesJanJun { get; set; }
            public string SalesJulDec { get; set; }
            public string[] SalesPersonArray { get; set; }
            public Report()
            {

            }
        }
    }
}





