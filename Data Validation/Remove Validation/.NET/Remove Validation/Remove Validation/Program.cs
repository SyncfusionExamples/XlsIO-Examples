using System.IO;
using Syncfusion.XlsIO;

namespace Remove_Validation
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"../../../Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Removing data validation from the worksheet
                worksheet.UsedRange.Clear(ExcelClearOptions.ClearDataValidations);

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath(@"../../../Output/Output.xlsx"));
            }
        }
    }
}
