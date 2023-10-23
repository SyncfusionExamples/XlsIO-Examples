using System;
using Syncfusion.XlsIO;
using System.IO;

namespace CSV_To_JSON
{
    class Program
    {
        static void Main(string[] args)
        {
            Conversions obj = new Conversions();
            obj.WorkbookToJSON();
            obj.RangeToJSON();
        }
    }
    public class Conversions
    {
        public void WorkbookToJSON()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.csv", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Active worksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Save the workbook to a JSON file
                workbook.SaveAsJson("WorkbookToJSON.json");
            }
        }
        public void RangeToJSON()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.csv", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                //Get the worksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Select a range
                IRange range = worksheet.Range["A2:A5"];

                //Save the range to a JSON file
                workbook.SaveAsJson("RangeToJSON.json", range);
            }
        }
    }
}
