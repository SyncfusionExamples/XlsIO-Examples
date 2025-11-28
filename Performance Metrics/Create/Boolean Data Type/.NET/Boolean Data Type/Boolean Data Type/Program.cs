using System.IO;
using Syncfusion.XlsIO;

namespace Boolean_Data_Type
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

                //Fill 100,000 rows × 50 columns with boolean
                for (int row = 1; row <= 100000; row++)
                {
                    for (int column = 1; column <= 50; column++)
                    {
                        //Boolean
                        sheet.SetBoolean(row, column, count % 2 == 0);

                        count++;
                    }

                }

                workbook.SaveAs(@"../../../Output/Output.xlsx");
            }
        }
    }
}





