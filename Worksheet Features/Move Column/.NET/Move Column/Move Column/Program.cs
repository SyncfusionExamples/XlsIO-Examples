using System.IO;
using Syncfusion.XlsIO;

namespace Move_Column
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream sourceStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(sourceStream);

                IWorksheet sourceWorksheet = workbook.Worksheets[0];
                IWorksheet destinationWorksheet = workbook.Worksheets[1];

                IRange source= sourceWorksheet.Range[1, 2];
                IRange destination = destinationWorksheet.Range[1, 2];

                //Move the Entire column to the next sheet
                source.EntireColumn.MoveTo(destination);

                //Saving the workbook
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
            }
        }
    }
}

