using Syncfusion.XlsIO;
using System.IO;
using static System.Net.WebRequestMethods;

namespace String_Data_Type
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];
                               
                int count = 0;
                // Fill 10,000 rows × 50 columns with string
                for (int row = 1; row <= 100000; row++)
                {
                    for (int column = 1; column <= 50; column++)
                    {

                        sheet.SetText(row, column, "One" + count);

                        count++;
                    }
                }
                workbook.SaveAs(@"../../../Output/Output.xlsx");
            }
        }
    }
}





