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
                IWorksheet sheet = workbook.Worksheets[0];

                //Fill 150 rows × 10,000 columns with number
                for (int row = 1; row <= 150; row++)
                {
                    for (int col = 1; col <= 10000; col++)
                    {
                        sheet[row, col].Number = row * col;
                    }
                }
            }
        }
    }
}





