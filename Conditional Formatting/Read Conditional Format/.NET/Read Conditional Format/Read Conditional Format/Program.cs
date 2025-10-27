using System.IO;
using Syncfusion.XlsIO;

namespace Read_Conditional_Format
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"), ExcelOpenType.Automatic);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Read conditional formatting settings 
                string formatType = worksheet.Range["A1"].ConditionalFormats[0].FormatType.ToString();
                string cfOperator = worksheet.Range["A1"].ConditionalFormats[0].Operator.ToString();

                workbook.SaveAs("Output.xlsx");
            }
        }
    }
}





