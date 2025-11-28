using System.IO;
using Syncfusion.XlsIO;

namespace Number_Data_Type
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

                //Access the first worksheet which contains data
                IWorksheet sheet = workbook.Worksheets[0];

                int count = 0;

                //Fill 100,000 rows × 50 columns with number
                for (int row = 1; row <= 100000; row++)
                {
                    for (int column = 1; column <= 50; column++)
                    {
                        //Number set method
                        sheet.SetNumber(row, column, count);

                        count++;
                    }
                }

                workbook.SaveAs(@"../../../Output/Output.xlsx");
            }
        }
    }
}





