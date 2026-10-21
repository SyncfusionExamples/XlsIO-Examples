using System.Data;
using DevExpress.Spreadsheet;

sealed class DevExpressEngine : IExcelEngine
{
    public string Name => "DevExpress Spreadsheet";

    public void Run(string scenario, string sourcePath, DataTable table, string outputStem)
    {
        using var workbook = new Workbook();
        switch (scenario)
        {
            case "Open and Save":
                workbook.LoadDocument(sourcePath);
                workbook.SaveDocument(outputStem + ".xlsx", DocumentFormat.Xlsx);
                break;
            case "Create and Save":
                WriteGrid(workbook.Worksheets[0], table);
                workbook.SaveDocument(outputStem + ".xlsx", DocumentFormat.Xlsx);
                break;
            case "Import DataTable":
                workbook.Worksheets[0].Import(table, true, 0, 0);
                workbook.SaveDocument(outputStem + ".xlsx", DocumentFormat.Xlsx);
                break;
            case "Export DataTable":
            {
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
                if (count < 1000)
                {
                    throw new InvalidOperationException($"Unexpected export row count {count}");
                }
                break;
            }
            case "Excel to CSV":
                workbook.LoadDocument(sourcePath);
                workbook.SaveDocument(outputStem + ".csv", DocumentFormat.Csv);
                break;
            case "Excel to JSON":
            {
                workbook.LoadDocument(sourcePath);
                Worksheet sheet = workbook.Worksheets[0];
                DevExpress.Spreadsheet.CellRange used = sheet.GetUsedRange();
                JsonExport.Write(outputStem + ".json", used.BottomRowIndex + 1, used.RightColumnIndex + 1, (r, c) => sheet.Cells[r - 1, c - 1].DisplayText);
                break;
            }
        }
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

