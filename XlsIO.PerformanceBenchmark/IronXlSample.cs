using System.Data;
using IronXL;

static class IronXlSample
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
        WorkBook workbook = WorkBook.Load(sourcePath);
        workbook.SaveAs(outputStem + ".xlsx");
    }

    public static void CreateAndSave(DataTable table, string outputStem)
    {
        WorkBook workbook = WorkBook.Create();
        WriteGrid(workbook.DefaultWorkSheet, table);
        workbook.SaveAs(outputStem + ".xlsx");
    }

    public static void ImportDataTable(DataTable table, string outputStem)
    {
        WorkBook workbook = WorkBook.Create();
        workbook.LoadWorkSheet(table);
        workbook.SaveAs(outputStem + ".xlsx");
    }

    public static void ExportDataTable(string sourcePath, int expectedRows)
    {
        WorkBook workbook = WorkBook.Load(sourcePath);
        DataTable exported = workbook.DefaultWorkSheet.ToDataTable(true);
        BenchmarkHost.EnsureRows(exported.Rows.Count, expectedRows);
    }

    public static void ExcelToCsv(string sourcePath, string outputStem)
    {
        WorkBook workbook = WorkBook.Load(sourcePath);
        workbook.SaveAsCsv(outputStem + ".csv");
    }

    public static void ExcelToJson(string sourcePath, string outputStem)
    {
        WorkBook workbook = WorkBook.Load(sourcePath);
        WorkSheet sheet = workbook.DefaultWorkSheet;
        JsonExport.Write(outputStem + ".json", sheet.RowCount, sheet.ColumnCount, (r, c) => sheet.GetCellAt(r - 1, c - 1)?.StringValue ?? "");
    }

    public static void ExcelToPdf(string sourcePath, string outputStem)
    {
        throw new NotSupportedException("IronXL.Excel 2026.10.5 does not include an Excel to PDF API.");
    }

    static void WriteGrid(WorkSheet sheet, DataTable table)
    {
        for (int c = 0; c < table.Columns.Count; c++)
        {
            sheet.SetCellValue(0, c, table.Columns[c].ColumnName);
            for (int r = 0; r < table.Rows.Count; r++)
            {
                sheet.SetCellValue(r + 1, c, table.Rows[r][c]);
            }
        }
    }
}
