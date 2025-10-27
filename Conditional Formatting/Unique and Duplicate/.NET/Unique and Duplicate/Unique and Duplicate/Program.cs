using System.IO;
using Syncfusion.XlsIO;
using Syncfusion.Drawing;

namespace Unique_and_Duplicate
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Fill worksheet with data
                worksheet.Range["A1:B1"].Merge();
                worksheet.Range["A1:B1"].CellStyle.Font.RGBColor = Color.FromArgb(255, 102, 102, 255);
                worksheet.Range["A1:B1"].CellStyle.Font.Size = 14;
                worksheet.Range["A1:B1"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                worksheet.Range["A1"].Text = "Global Internet Usage";
                worksheet.Range["A1:B1"].CellStyle.Font.Bold = true;

                worksheet.Range["A3:B21"].CellStyle.Font.RGBColor = Color.FromArgb(255, 64, 64, 64);
                worksheet.Range["A3:B3"].CellStyle.Font.Bold = true;
                worksheet.Range["B3"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignRight;

                worksheet.Range["A3"].Text = "Country";
                worksheet.Range["A4"].Text = "Northern America";
                worksheet.Range["A5"].Text = "Central America";
                worksheet.Range["A6"].Text = "The Caribbean";
                worksheet.Range["A7"].Text = "South America";
                worksheet.Range["A8"].Text = "Northern Europe";
                worksheet.Range["A9"].Text = "Eastern Europe";
                worksheet.Range["A10"].Text = "Western Europe";
                worksheet.Range["A11"].Text = "Southern Europe";
                worksheet.Range["A12"].Text = "Northern Africa";
                worksheet.Range["A13"].Text = "Eastern Africa";
                worksheet.Range["A14"].Text = "Middle Africa";
                worksheet.Range["A15"].Text = "Western Africa";
                worksheet.Range["A16"].Text = "Southern Africa";
                worksheet.Range["A17"].Text = "Central Asia";
                worksheet.Range["A18"].Text = "Eastern Asia";
                worksheet.Range["A19"].Text = "Southern Asia";
                worksheet.Range["A20"].Text = "SouthEast Asia";
                worksheet.Range["A21"].Text = "Oceania";

                worksheet.Range["B3"].Text = "Usage";
                worksheet.SetValue(4, 2, "88%");
                worksheet.SetValue(5, 2, "61%");
                worksheet.SetValue(6, 2, "49%");
                worksheet.SetValue(7, 2, "68%");
                worksheet.SetValue(8, 2, "94%");
                worksheet.SetValue(9, 2, "74%");
                worksheet.SetValue(10, 2, "90%");
                worksheet.SetValue(11, 2, "77%");
                worksheet.SetValue(12, 2, "49%");
                worksheet.SetValue(13, 2, "27%");
                worksheet.SetValue(14, 2, "12%");
                worksheet.SetValue(15, 2, "39%");
                worksheet.SetValue(16, 2, "51%");
                worksheet.SetValue(17, 2, "50%");
                worksheet.SetValue(18, 2, "58%");
                worksheet.SetValue(19, 2, "36%");
                worksheet.SetValue(20, 2, "58%");
                worksheet.SetValue(21, 2, "69%");

                worksheet.SetColumnWidth(1, 23.45);
                worksheet.SetColumnWidth(2, 8.09);

                IConditionalFormats conditionalFormats =
                worksheet.Range["A4:B21"].ConditionalFormats;
                IConditionalFormat condition = conditionalFormats.AddCondition();

                //conditional format to set duplicate format type
                condition.FormatType = ExcelCFType.Duplicate;
                condition.BackColorRGB = Color.FromArgb(255, 255, 199, 206);

                #region Save
                //Saving the workbook
                workbook.SaveAs(Path.GetFullPath("Output/UniqueandDuplicate.xlsx"));
                #endregion
            }
        }
    }
}




