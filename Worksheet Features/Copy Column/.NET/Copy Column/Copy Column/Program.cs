using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Interfaces;

namespace Copy_Column
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

                IRange sourceColumn = sourceWorksheet.Range[1, 1];
                IRange destinationColumn = destinationWorksheet.Range[1, 1];

                //Copy the entire column to the next sheet
                sourceColumn.EntireColumn.CopyTo(destinationColumn);

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
            }
        }
    }
}




