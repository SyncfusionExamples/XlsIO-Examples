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
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"../../../Data/Input.xlsx"));

                //Access the first worksheet which contains data
                IWorksheet sheet = workbook.Worksheets[0];

                IRange usedRange = sheet.UsedRange;

                int firstRow = usedRange.Row;
                int lastRow = usedRange.LastRow;
                int firstcolumn = usedRange.Column;
                int lastcolumn = usedRange.LastColumn;
                for (int row = firstRow; row <= lastRow; row++)
                {
                    for (int column = firstcolumn; column <= lastcolumn; column++)
                    {
                        var value = sheet.GetCellValue(row, column, false);
                    }
                }
            }
        }
    }
}





