using Syncfusion.XlsIO;

namespace Set_Cell_Values
{
    class Program
    {
        static void Main()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Open(Path.GetFullPath(@"Data/Input.xlsx"));
                IWorksheet sheet = workbook.Worksheets[0];

                sheet.Range["A4"].Text = "Gamma";
                sheet.Range["B4"].Number = 3200;
                sheet.Range["C4"].DateTime = new DateTime(2026, 7, 25);
                sheet.Range["D4"].Boolean = true;

                sheet.Range["A5"].Text = "Delta";
                sheet.Range["B5"].Number = 4500;
                sheet.Range["C5"].DateTime = new DateTime(2026, 8, 1);
                sheet.Range["D5"].Boolean = false;

                sheet.Range["B6"].Formula = "SUM(B2:B5)";

                sheet.Range["A1:D1"].CellStyle.Font.Bold = true;
                sheet.Range["A1:D6"].CellStyle.HorizontalAlignment = ExcelHAlign.HAlignCenter;

                workbook.SaveAs(Path.GetFullPath("Output/Output.xlsx"));
                workbook.Close();
            }
        }
    }
}
