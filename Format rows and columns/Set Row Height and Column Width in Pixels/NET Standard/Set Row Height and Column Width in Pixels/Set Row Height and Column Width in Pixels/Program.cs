using Syncfusion.XlsIO;
namespace Set_Row_Height_and_Column_width_in_pixels
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Load an existing file
                FileStream inputStream = new FileStream("../../../Data/InputTemplate.xlsx", FileMode.Open);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set row height in pixels
                worksheet.SetRowHeightInPixels(2, 50);

                //Get column width in pixels
                worksheet.SetColumnWidthInPixels(3, 100);

                //Save the workbook as stream
                FileStream outputStream = new FileStream("Output.xlsx", FileMode.Create, FileAccess.Write);
                workbook.SaveAs(outputStream);

                //Dispose stream
                inputStream.Dispose();
                outputStream.Dispose();
            }
        }
    }
}