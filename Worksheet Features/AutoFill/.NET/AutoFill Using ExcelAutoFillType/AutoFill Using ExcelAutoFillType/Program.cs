using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;

namespace AutoFill_Using_ExcelAutoFillType
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Setting the values to the cells
                worksheet["A1"].Number = 1;
                worksheet["A2"].Number = 3;
                worksheet["A3"].Number = 5;

                //Define the source range
                IRange source = worksheet["A1:A3"];

                //Define the destination range
                IRange destinationRange = worksheet["A4:A10"];

                //Use AutoFill method to fill the values based on ExcelAutoFillType
                source.AutoFill(destinationRange, ExcelAutoFillType.FillSeries);

                //Saving the workbook as stream
                FileStream outputStream = new FileStream(Path.GetFullPath("Output/Output.xlsx"), FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose streams
                outputStream.Dispose();
            }

        }
    }
}
