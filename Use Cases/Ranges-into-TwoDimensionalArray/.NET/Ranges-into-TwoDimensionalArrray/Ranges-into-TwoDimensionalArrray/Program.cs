using Syncfusion.XlsIO;

namespace Ranges_into_TwoDimensionalArray
{
    class Program
    {
        public static void Main(string[] args)
        {
            //Create an instance of ExcelEngine
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                //Instantiate the Excel application object
                IApplication application = excelEngine.Excel;

                //Set the default application version as Excel 2016
                application.DefaultVersion = ExcelVersion.Xlsx;

                //Loads an existing workbook
                FileStream fileStream = new FileStream(@"../../../Data/InputTemplate.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set the range to convert into two dimensional array
                IRange range = worksheet["B1:E5"];
                string[,] arrayOfArrays = ConvertIRangeToArray(worksheet, range);
            }
        }
        /// <summary>
        /// Converting range values into 2 dimensional array
        /// </summary>
        public static string[,] ConvertIRangeToArray(IWorksheet worksheet, IRange range)
        {
            int startRow = range.Row;
            int startCol = range.Column;
            int endRow = range.LastRow;
            int endCol = range.LastColumn;

            string[,] numbers = new string[endRow - startRow + 1, endCol - startCol + 1];

            for (int i = 0; i <= endRow - startRow; i++)
            {
                for (int j = 0; j <= endCol - startCol; j++)
                {
                    numbers[i, j] = worksheet[startRow + i, startCol + j].Value;
                    Console.Write(numbers[i, j]);
                    Console.Write("\t");
                }
                Console.Write("\n______________________________________________\n");
            }
            return numbers;
        }
    }
}