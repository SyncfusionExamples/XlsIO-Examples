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
                FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Data/Input.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet sheet = workbook.Worksheets[0];

                //Read 150 rows × 10,000 columns
                for (int row = 1; row <= 150; row++)
                {
                    for (int col = 1; col <= 10000; col++)
                    {
                        bool boolVal = sheet[row, col].Boolean;
                    }
                }

                //Dispose streams
                fileStream.Dispose();
            }
        }
    }
}





