using System.IO;
using Syncfusion.XlsIO;

namespace Access_Worksheet
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../../InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                #region  Access
                //Accessing via index
                IWorksheet sheet = workbook.Worksheets[0];

                //Accessing via sheet name
                IWorksheet NamedSheet = workbook.Worksheets["Sample"];
                #endregion

                //Dispose streams
                inputStream.Dispose();
            }
        }
    }
}
