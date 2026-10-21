using System.Data;
using GemBox.Spreadsheet;

sealed class GemBoxEngine : IExcelEngine
{
    public string Name => "GemBox.Spreadsheet";

    public void Run(string scenario, string sourcePath, DataTable table, string outputStem)
    {
        switch (scenario)
        {
            case "Open and Save":
            {
                ExcelFile workbook = ExcelFile.Load(sourcePath);
                workbook.GroupMethodsAffectedCellsLimit = int.MaxValue;
                workbook.Save(outputStem + ".xlsx");
                break;
            }
            case "Create and Save":
            {
                var workbook = new ExcelFile();
                ExcelWorksheet sheet = workbook.Worksheets.Add("Sales");
                WriteGrid(sheet, table);
                workbook.Save(outputStem + ".xlsx");
                break;
            }
            case "Import DataTable":
            {
                var workbook = new ExcelFile();
                ExcelWorksheet sheet = workbook.Worksheets.Add("Sales");
                sheet.InsertDataTable(table, new InsertDataTableOptions { ColumnHeaders = true, StartRow = 0, StartColumn = 0 });
                workbook.Save(outputStem + ".xlsx");
                break;
            }
            case "Export DataTable":
            {
                ExcelFile workbook = ExcelFile.Load(sourcePath);
                DataTable exported = workbook.Worksheets[0].CreateDataTable(new CreateDataTableOptions { ColumnHeaders = true, StartRow = 0, StartColumn = 0 });
                if (exported.Rows.Count < 1000)
                {
                    throw new InvalidOperationException($"Unexpected export row count {exported.Rows.Count}");
                }
                break;
            }
            case "Excel to CSV":
            {
                ExcelFile workbook = ExcelFile.Load(sourcePath);
                workbook.Save(outputStem + ".csv");
                break;
            }
            case "Excel to JSON":
            {
                ExcelFile workbook = ExcelFile.Load(sourcePath);
                ExcelWorksheet sheet = workbook.Worksheets[0];
                int rows = sheet.Rows.Count;
                int cols = sheet.CalculateMaxUsedColumns();
                JsonExport.Write(outputStem + ".json", rows, cols, (r, c) => sheet.Cells[r - 1, c - 1].Value?.ToString() ?? "");
                break;
            }
        }
    }

    static void WriteGrid(ExcelWorksheet sheet, DataTable table)
    {
        for (int c = 0; c < table.Columns.Count; c++)
        {
            sheet.Cells[0, c].Value = table.Columns[c].ColumnName;
            var column = new object[table.Rows.Count];
            for (int r = 0; r < table.Rows.Count; r++)
            {
                column[r] = table.Rows[r][c];
            }

            sheet.Cells.GetSubrangeAbsolute(1, c, table.Rows.Count, c).Value = column;
        }
    }
}

sealed record BenchResult(
    string Library,
    string Scenario,
    string Status,
    double MedianMs,
    double PeakWorkingSetMb,
    double DeltaWorkingSetMb,
    double AllocatedMb,
    string? Error);
