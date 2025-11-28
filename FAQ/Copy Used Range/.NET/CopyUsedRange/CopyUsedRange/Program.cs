using Syncfusion.XlsIO;

namespace CopyUsedRange
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook sourceWorkbook = excelEngine.Excel.Workbooks.Open(Path.GetFullPath(@"Data/Source.xlsx"));
                IWorkbook destinationWorkbook = excelEngine.Excel.Workbooks.Open(Path.GetFullPath(@"Data/Destination.xlsx"));

                IWorksheet sourceWorksheet = sourceWorkbook.Worksheets["Sheet1"];
                IWorksheet destinationWorksheet = destinationWorkbook.Worksheets["Sheet1"];

                //Get the actual used range from source sheet
                IRange sourceRange = sourceWorksheet.UsedRange;

                //Copy the entire used range from source sheet to destination sheet
                sourceRange.CopyTo(destinationWorksheet.Range[sourceRange.Row, sourceRange.Column]);

                //Save the destination workbook
                destinationWorkbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }
        }
    }
}