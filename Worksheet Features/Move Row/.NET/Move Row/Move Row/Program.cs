using System.IO;
using Syncfusion.XlsIO;

namespace Move_Row
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

                IWorksheet sourceWorksheet = workbook.Worksheets[0];
                IWorksheet destinationWorksheet = workbook.Worksheets[1];

                IRange sourceRow = sourceWorksheet.Range[2, 1];
                IRange destinationRow = destinationWorksheet.Range[2, 1];

                //Move the Entire row to the next sheet
                sourceRow.EntireRow.MoveTo(destinationRow);

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}





