using System.Data;
using DevExpress.Spreadsheet;

static class DevExpressSpreadsheet
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
        using var workbook = new Workbook();
        workbook.LoadDocument(sourcePath);
        workbook.SaveDocument(outputStem + ".xlsx", DocumentFormat.Xlsx);
    }

    public static void CreateAndSave(DataTable table, string outputStem)
    {
        using var workbook = new Workbook();
        WriteGrid(workbook.Worksheets[0], table);
        workbook.SaveDocument(outputStem + ".xlsx", DocumentFormat.Xlsx);
    }

    public static void ImportDataTable(DataTable table, string outputStem)
    {
        using var workbook = new Workbook();
        workbook.Worksheets[0].Import(table, true, 0, 0);
        workbook.SaveDocument(outputStem + ".xlsx", DocumentFormat.Xlsx);
    }

    public static void ExportDataTable(string sourcePath, int expectedRows)
    {
        using var workbook = new Workbook();
        workbook.LoadDocument(sourcePath);
        Worksheet sheet = workbook.Worksheets[0];
        CellRange used = sheet.GetUsedRange();
        object dataSource = used.GetDataSource();
        int count = dataSource switch
        {
            DataTable dataTable => dataTable.Rows.Count,
            System.Collections.ICollection collection => collection.Count,
            System.Collections.IEnumerable enumerable => enumerable.Cast<object>().Count(),
            _ => throw new InvalidOperationException($"Unexpected data source {dataSource?.GetType().FullName}"),
        };
        BenchmarkHost.EnsureRows(count, expectedRows);
    }

    public static void ExcelToCsv(string sourcePath, string outputStem)
    {
        using var workbook = new Workbook();
        workbook.LoadDocument(sourcePath);
        workbook.SaveDocument(outputStem + ".csv", DocumentFormat.Csv);
    }

    public static void ExcelToJson(string sourcePath, string outputStem)
    {
        using var workbook = new Workbook();
        workbook.LoadDocument(sourcePath);
        Worksheet sheet = workbook.Worksheets[0];
        CellRange used = sheet.GetUsedRange();
        JsonExport.Write(outputStem + ".json", used.BottomRowIndex + 1, used.RightColumnIndex + 1, (r, c) => sheet.Cells[r - 1, c - 1].DisplayText);
    }

    public static void ExcelToPdf(string sourcePath, string outputStem)
    {
        using var workbook = new Workbook();
        workbook.LoadDocument(sourcePath);
        workbook.ExportToPdf(outputStem + ".pdf");
    }

    static void WriteGrid(Worksheet sheet, DataTable table)
    {
        var values = new object[table.Rows.Count + 1, table.Columns.Count];
        for (int c = 0; c < table.Columns.Count; c++)
        {
            values[0, c] = table.Columns[c].ColumnName;
        }

        for (int r = 0; r < table.Rows.Count; r++)
        {
            for (int c = 0; c < table.Columns.Count; c++)
            {
                values[r + 1, c] = table.Rows[r][c];
            }
        }

        sheet.Import(values, 0, 0);
    }
}
