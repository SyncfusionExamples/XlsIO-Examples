using Syncfusion.XlsIO;


class Program
{
    static void Main(string[] args)
    {
        // Initialize Excel
        ExcelEngine excelEngine = new ExcelEngine();
        IApplication application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;
        IWorkbook workbook = application.Workbooks.Create(1);
        IWorksheet sheet = workbook.Worksheets[0];
        sheet.Name = "CF Formula Offset";

        // Headers
        sheet["L22"].Text = "Value L"; sheet["M22"].Text = "Status M";
        sheet["P22"].Text = "Value P"; sheet["Q22"].Text = "Status Q";
        sheet["AR22"].Text = "Value AR";
        sheet["L22:AR22"].CellStyle.Font.Bold = true;

        // Sample data (rows 23-32)
        int[] valuesL = { 150, -50, 200, 0, -100, 75, -25, 300, 0, 125 };
        string[] statusM = { "ok", "n.m.", "n.m.", "pending", "n.m.", "ok", "n.m.", "complete", "n.m.", "ok" };
        int[] valuesP = { 80, 120, -40, 95, -60, 200, 0, -15, 170, 85 };
        string[] statusQ = { "n.m.", "ok", "n.m.", "n.m.", "ok", "n.m.", "pending", "n.m.", "ok", "n.m." };
        int[] valuesAR = { 45, -30, 90, 0, -75, 110, 25, -10, 50, 135 };

        for (int i = 0; i < 10; i++)
        {
            sheet[23 + i, 12].Number = valuesL[i];    // Column L (12)
            sheet[23 + i, 13].Text = statusM[i];      // Column M (13)
            sheet[23 + i, 16].Number = valuesP[i];    // Column P (16)
            sheet[23 + i, 17].Text = statusQ[i];      // Column Q (17)
            sheet[23 + i, 44].Number = valuesAR[i];   // Column AR (44)
        }

        // Apply conditional formatting rules to columns L, P, AR
        string[] targetCols = { "L", "P", "AR" };
        string[] adjacentCols = { "M", "Q", "AS" }; // Columns to check for Rule 2

        for (int c = 0; c < targetCols.Length; c++)
        {
            string range = $"{targetCols[c]}23:{targetCols[c]}32";
            IConditionalFormats formats = sheet[range].ConditionalFormats;

            // Rule 1: Value > 0 → GREEN (Offset: 0,0)
            IConditionalFormat rule1 = formats.AddCondition();
            rule1.FormatType = ExcelCFType.Formula;
            rule1.FirstFormula = $"={targetCols[c]}23>0";
            rule1.BackColor = ExcelKnownColors.Light_green;

            // Rule 2: Adjacent column = "n.m." → YELLOW (Offset: 0,1)
            IConditionalFormat rule2 = formats.AddCondition();
            rule2.FormatType = ExcelCFType.Formula;
            rule2.FirstFormula = $"={adjacentCols[c]}23=\"n.m.\"";
            rule2.BackColor = ExcelKnownColors.Light_yellow;

            // Rule 3: Value < 0 → ORANGE (Offset: 0,0)
            IConditionalFormat rule3 = formats.AddCondition();
            rule3.FormatType = ExcelCFType.Formula;
            rule3.FirstFormula = $"={targetCols[c]}23<0";
            rule3.BackColor = ExcelKnownColors.Light_orange;
            rule3.FontColor = ExcelKnownColors.Red;
        }

        // Calculate and display offset for specific cells
        Console.WriteLine("\n=== OFFSET CALCULATIONS ===\n");
        PrintOffset("L23", "L23");   // Rule 1: Same cell
        PrintOffset("M23", "L23");   // Rule 2: One column right
        PrintOffset("P25", "P25");   // Rule 1: Same cell
        PrintOffset("Q26", "P26");   // Rule 2: One column right
        PrintOffset("AS28", "AR28"); // Rule 2: One column right
        PrintOffset("L30", "L25");   // Different row

        Console.WriteLine("\n=== OFFSET FORMULA ===");
        Console.WriteLine("Row Offset = Formula Row - Applied Row");
        Console.WriteLine("Col Offset = Formula Col - Applied Col");

        // Save file
        workbook.SaveAs(Path.GetFullPath("Output/CF_FormulaOffset.xlsx"));
        workbook.Close();
        excelEngine.Dispose();

        Console.WriteLine("\nFile created: CF_FormulaOffset.xlsx");
    }

    // Calculate and print offset between formula cell and applied cell
    static void PrintOffset(string formulaCell, string appliedCell)
    {
        int formRow = int.Parse(new string(formulaCell.Where(char.IsDigit).ToArray()));
        int applRow = int.Parse(new string(appliedCell.Where(char.IsDigit).ToArray()));

        string formCol = new string(formulaCell.Where(char.IsLetter).ToArray());
        string applCol = new string(appliedCell.Where(char.IsLetter).ToArray());

        int formColNum = ColToNum(formCol);
        int applColNum = ColToNum(applCol);

        int rowOffset = formRow - applRow;
        int colOffset = formColNum - applColNum;

        Console.WriteLine($"Formula={formulaCell}, Applied={appliedCell}, Offset({rowOffset},{colOffset})");
    }

    // Convert column letter to number (A=1, B=2, AR=44, AS=45, etc.)
    static int ColToNum(string col)
    {
        int result = 0;
        for (int i = 0; i < col.Length; i++)
            result = result * 26 + (col[i] - 'A' + 1);
        return result;
    }
}
