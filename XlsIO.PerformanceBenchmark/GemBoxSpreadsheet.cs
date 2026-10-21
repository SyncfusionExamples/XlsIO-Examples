using System.Data;
using GemBox.Spreadsheet;

static class GemBoxSpreadsheet
{
    public static void Run(string method, string sourcePath, DataTable table, string outputStem)
    {
        switch (method)
        {
            case "OpenAndSave": OpenAndSave(sourcePath, outputStem); break;
            case "CreateAndSave": CreateAndSave(table, outputStem); break;
            case "ImportDataTable": ImportDataTable(table, outputStem); break;
            case "ExportDataTable": ExportDataTable(sourcePath, table.Rows.Count); break;
            case "ExcelToCsv": ExcelToCsv(sourcePath, outputStem); break;
            case "ExcelToJson": ExcelToJson(sourcePath, outputStem); break;
            case "ExcelToPdf": ExcelToPdf(sourcePath, outputStem); break;
            default: throw new ArgumentException($"Unknown method '{method}'.");
        }
    }

    public static void OpenAndSave(string sourcePath, string outputStem)
    {
        ExcelFile workbook = ExcelFile.Load(sourcePath);
        workbook.Save(outputStem + ".xlsx");
    }

    public static void CreateAndSave(DataTable table, string outputStem)
    {
        var workbook = new ExcelFile();
        ExcelWorksheet sheet = workbook.Worksheets.Add("Sales");
        WriteGrid(sheet, table);
        workbook.Save(outputStem + ".xlsx");
    }

    public static void ImportDataTable(DataTable table, string outputStem)
    {
        var workbook = new ExcelFile();
        ExcelWorksheet sheet = workbook.Worksheets.Add("Sales");
        sheet.InsertDataTable(table, new InsertDataTableOptions { ColumnHeaders = true, StartRow = 0, StartColumn = 0 });
        workbook.Save(outputStem + ".xlsx");
    }

    public static void ExportDataTable(string sourcePath, int expectedRows)
    {
        ExcelFile workbook = ExcelFile.Load(sourcePath);
        DataTable exported = workbook.Worksheets[0].CreateDataTable(new CreateDataTableOptions { ColumnHeaders = true, StartRow = 0, StartColumn = 0 });
        BenchmarkHost.EnsureRows(exported.Rows.Count, expectedRows);
    }

    public static void ExcelToCsv(string sourcePath, string outputStem)
    {
        ExcelFile workbook = ExcelFile.Load(sourcePath);
        workbook.Save(outputStem + ".csv");
    }

    public static void ExcelToJson(string sourcePath, string outputStem)
    {
        ExcelFile workbook = ExcelFile.Load(sourcePath);
        ExcelWorksheet sheet = workbook.Worksheets[0];
        int rows = sheet.Rows.Count;
        int cols = sheet.CalculateMaxUsedColumns();
        JsonExport.Write(outputStem + ".json", rows, cols, (r, c) => sheet.Cells[r - 1, c - 1].Value?.ToString() ?? "");
    }

    public static void ExcelToPdf(string sourcePath, string outputStem)
    {
        ExcelFile workbook = ExcelFile.Load(sourcePath);
        workbook.Save(outputStem + ".pdf");
    }

    static void WriteGrid(ExcelWorksheet sheet, DataTable table)
    {
        for (int c = 0; c < table.Columns.Count; c++)
        {
            sheet.Cells[0, c].Value = table.Columns[c].ColumnName;
            for (int r = 0; r < table.Rows.Count; r++)
            {
                sheet.Cells[r + 1, c].Value = table.Rows[r][c];
            }
        }
    }
}
