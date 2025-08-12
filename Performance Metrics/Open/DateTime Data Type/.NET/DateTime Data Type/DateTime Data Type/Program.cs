using System.IO;
using Syncfusion.XlsIO;

namespace DateTime_Data_Type
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

                //Dispose streams
                fileStream.Dispose();
            }
        }
    }
}





