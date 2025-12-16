using Syncfusion.XlsIO;

namespace FirstUsedCellInUsedRange
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the used range of the worksheet
                IRange usedRange = worksheet.UsedRange;

                //Get the first cell from the used range
                IRange firstCell = worksheet.Range[usedRange.Row, usedRange.Column];

                //Get the address of the first cell
                string firstCellAddress = firstCell.AddressLocal;

                //Display the address of the first cell
                Console.WriteLine("The address of the first used cell in used range is: " + firstCellAddress);

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }
        }
    }
}