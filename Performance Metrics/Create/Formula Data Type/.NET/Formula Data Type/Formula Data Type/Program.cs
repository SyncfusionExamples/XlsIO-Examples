using System.IO;
using Syncfusion.XlsIO;

namespace Formula_Data_Type
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

                //Fill 100,000 rows × 50 columns with formula
                for (int row = 2; row <= 100000; row++)
                {
                    for (int column = 1; column <= 50; column++)
                    {
                        sheet.SetFormula(row, column, "IF(A1>10,SUM(A1,10),10)");
                    }
                }

                workbook.SaveAs(@"../../../Output/Output.xlsx");
            }
        }
    }
}





