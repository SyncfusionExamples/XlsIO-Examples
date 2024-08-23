using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.Collections;
using Syncfusion.XlsIO.Implementation;
using Syncfusion.XlsIO.Interfaces;
using static System.Net.Mime.MediaTypeNames;

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

                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);

                IWorksheet sourceWorksheet = workbook.Worksheets[0];
                IWorksheet destinationWorksheet = workbook.Worksheets[1];

                IRange sourceColumn = sourceWorksheet.Range[1, 1];
                IRange destinationColumn = destinationWorksheet.Range[1, 1];

                //Copy the entire column to the next sheet
                sourceColumn.EntireColumn.CopyTo(destinationColumn);

                //Saving the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
                inputStream.Dispose();
            }
        }
    }
}
