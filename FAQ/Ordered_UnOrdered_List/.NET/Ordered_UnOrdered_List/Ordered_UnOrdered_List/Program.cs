using Syncfusion.XlsIO;

class Program
{
    static void Main(string[] args)
    {
        using (ExcelEngine excelEngine = new ExcelEngine())
        {
            IApplication application = excelEngine.Excel;
            application.DefaultVersion = ExcelVersion.Xlsx;

            IWorkbook workbook = application.Workbooks.Create(1);
            IWorksheet sheet = workbook.Worksheets[0];

            // Target cell
            IRange cell = sheet.Range["B3"];
            IRichTextString richText = cell.RichText;

            // Add initial text
            richText.Text = "list:\n";

            // Fonts
            IFont bulletFont = workbook.CreateFont();
            bulletFont.FontName = "Courier New";
            bulletFont.Size = 10;

            IFont textFont = workbook.CreateFont();
            textFont.FontName = "Segoe UI";
            textFont.Size = 10;

            // First bullet
            richText.Text += "  • number1\n";
            richText.SetFont(7, 9, bulletFont);   // bullet
            richText.SetFont(10, 16, textFont);   // text

            // Second bullet
            richText.Text += "  • number2\n";
            richText.SetFont(18, 20, bulletFont); // bullet
            richText.SetFont(21, 27, textFont);   // text

            // Third bullet
            richText.Text += "  • number3";
            richText.SetFont(29, 31, bulletFont); // bullet
            richText.SetFont(32, 38, textFont);   // text

            // Wrap text so bullets appear on separate lines
            cell.CellStyle.WrapText = true;

            // Save the workbook
            workbook.SaveAs(Path.GetFullPath("Output/BulletedCell.xlsx"));
        }
    }
}
