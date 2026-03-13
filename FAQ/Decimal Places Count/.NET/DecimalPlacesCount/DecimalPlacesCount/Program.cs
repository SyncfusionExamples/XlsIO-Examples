using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation.PivotTables;

namespace DecimalPlacesCount
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/InputTemplate.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                // Get the cell text safely
                string cellText = worksheet.Range["G2"].Value?.ToString() ?? string.Empty;

                // Count decimal places: if there's a decimal point, count chars after it; otherwise 0
                int countDecimalPlaces = 0;
                int dotIndex = cellText.IndexOf('.');
                if (dotIndex >= 0)
                {
                    countDecimalPlaces = cellText.Length - dotIndex - 1;
                }

                // Display result in console
                Console.WriteLine(countDecimalPlaces);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                #endregion
            }
        }
    }
}