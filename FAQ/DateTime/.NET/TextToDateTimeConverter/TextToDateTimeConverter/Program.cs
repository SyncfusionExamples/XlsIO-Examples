using Syncfusion.XlsIO;
using System.Globalization;

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

                //Open the input workbook
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));

                //Access the first worksheet
                IWorksheet worksheet = workbook.Worksheets[0];

                //Get the used range to iterate through populated cells
                IRange used = worksheet.UsedRange;

                //Set culture and parsing styles for interpreting text dates
                CultureInfo culture = new CultureInfo("en-IN");
                DateTimeStyles styles = DateTimeStyles.None;

                //Iterate through the used range and convert text-formatted dates to DateTime
                for (int row = used.Row; row <= used.LastRow; row++)
                {
                    for (int col = used.Column; col <= used.LastColumn; col++)
                    {
                        IRange cell = worksheet[row, col];
                        DateTime date;

                        //Log if the cell already contains a true DateTime
                        if (cell.HasDateTime)
                        {
                            Console.WriteLine(cell.DateTime);
                        }
                        //Try parsing text using the specified culture and assign DateTime back to the cell
                        else if (DateTime.TryParse(cell.Value, culture, styles, out date))
                        {
                            cell.DateTime = date;
                        }
                    }
                }

                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath(@"Output/Output.xlsx"));
            }

        }
    }
}