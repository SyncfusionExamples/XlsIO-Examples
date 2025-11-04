using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;


namespace RGBValueCellColor
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

                //Apply cell color
                worksheet.Range["A1"].CellStyle.ColorIndex = ExcelKnownColors.Custom50;

                //Get the RGB values of the cell color
                Color color = worksheet.Range["A1"].CellStyle.Color;
                byte red = color.R;
                byte green = color.G;
                byte blue = color.B;

                //Print the RGB values
                Console.WriteLine($"Red: {red}, Green: {green}, Blue: {blue}");

                //Save the workbook
                workbook.SaveAs(Path.GetFullPath("../../../Output/Output.xlsx"));
            }
        }
    }
}
