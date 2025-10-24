using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using Syncfusion.XlsIO;

namespace Worksheet_to_CollectionObjects
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
                IWorksheet worksheet = workbook.Worksheets[0];

                //Export worksheet data into Collection Objects
                List<Report> collectionObjects = worksheet.ExportData<Report>(1, 1, 10, 3);
            }
        }
        public class Report
        {
            [DisplayNameAttribute("Sales Person Name")]
            public string SalesPerson { get; set; }
            [Bindable(false)]
            public string SalesJanJun { get; set; }
            public string SalesJulDec { get; set; }

            public Report()
            {

            }
        }
    }
}





