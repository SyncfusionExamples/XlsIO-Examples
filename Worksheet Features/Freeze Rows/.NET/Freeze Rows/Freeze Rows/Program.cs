using System.IO;
using Syncfusion.XlsIO;

namespace Freeze_Rows
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Applying freeze rows to the sheet by specifying a cell
                worksheet.Range["A3"].FreezePanes();

                //Set first visible row in the bottom pane
                worksheet.FirstVisibleRow = 3;

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}





