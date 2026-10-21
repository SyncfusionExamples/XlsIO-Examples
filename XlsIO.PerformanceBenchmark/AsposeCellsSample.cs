using System.Data;
using AsposeCells = Aspose.Cells;

static class AsposeCellsSample
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
        var workbook = new AsposeCells.Workbook(sourcePath);
        workbook.Save(outputStem + ".xlsx");
    }

    public static void CreateAndSave(DataTable table, string outputStem)
    {
        var workbook = new AsposeCells.Workbook();
        WriteGrid(workbook.Worksheets[0], table);
        workbook.Save(outputStem + ".xlsx");
    }

    public static void ImportDataTable(DataTable table, string outputStem)
    {
        var workbook = new AsposeCells.Workbook();
        workbook.Worksheets[0].Cells.ImportData(table, 0, 0, new AsposeCells.ImportTableOptions { IsFieldNameShown = true });
        workbook.Save(outputStem + ".xlsx");
    }

    public static void ExportDataTable(string sourcePath, int expectedRows)
    {
        var workbook = new AsposeCells.Workbook(sourcePath);
        AsposeCells.Cells cells = workbook.Worksheets[0].Cells;
        DataTable exported = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
        BenchmarkHost.EnsureRows(exported.Rows.Count, expectedRows);
    }

    public static void ExcelToCsv(string sourcePath, string outputStem)
    {
        var workbook = new AsposeCells.Workbook(sourcePath);
        workbook.Save(outputStem + ".csv", AsposeCells.SaveFormat.Csv);
    }

    public static void ExcelToJson(string sourcePath, string outputStem)
    {
        var workbook = new AsposeCells.Workbook(sourcePath);
        workbook.Save(outputStem + ".json", AsposeCells.SaveFormat.Json);
    }

    public static void ExcelToPdf(string sourcePath, string outputStem)
    {
        var workbook = new AsposeCells.Workbook(sourcePath);
        workbook.Save(outputStem + ".pdf", AsposeCells.SaveFormat.Pdf);
    }

    static void WriteGrid(AsposeCells.Worksheet sheet, DataTable table)
    {
        int columns = table.Columns.Count;
        var values = new object[table.Rows.Count + 1, columns];
        for (int c = 0; c < columns; c++)
        {
            values[0, c] = table.Columns[c].ColumnName;
        }

        for (int r = 0; r < table.Rows.Count; r++)
        {
            for (int c = 0; c < columns; c++)
            {
                values[r + 1, c] = table.Rows[r][c];
            }
        }

        sheet.Cells.ImportTwoDimensionArray(values, 0, 0);
    }
}
