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
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Export worksheet data into Collection Objects
                List<Report> collectionObjects = worksheet.ExportData<Report>(1, 1, 10, 3);

                //Dispose streams
                inputStream.Dispose();
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





