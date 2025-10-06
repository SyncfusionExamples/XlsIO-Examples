using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Interfaces;

namespace Copy_Row
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));

                IWorksheet sourceWorksheet = workbook.Worksheets[0];
                IWorksheet destinationWorksheet = workbook.Worksheets[1];

                IRange sourceRow = sourceWorksheet.Range[1, 1];
                IRange destinationRow = destinationWorksheet.Range[1, 1];

                //Copy the entire row to the next sheet
                sourceRow.EntireRow.CopyTo(destinationRow);

                //Saving the workbook as stream
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}




