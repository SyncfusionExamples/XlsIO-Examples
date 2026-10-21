using System.Data;
using AsposeCells = Aspose.Cells;

sealed class AsposeEngine : IExcelEngine
{
    public string Name => "Aspose.Cells";

    public void Run(string scenario, string sourcePath, DataTable table, string outputStem)
    {
        switch (scenario)
        {
            case "Open and Save":
            {
                var workbook = new AsposeCells.Workbook(sourcePath);
                workbook.Save(outputStem + ".xlsx");
                break;
            }
            case "Create and Save":
            {
                var workbook = new AsposeCells.Workbook();
                AsposeCells.Worksheet sheet = workbook.Worksheets[0];
                WriteGrid(sheet, table);
                workbook.Save(outputStem + ".xlsx");
                break;
            }
            case "Import DataTable":
            {
                var workbook = new AsposeCells.Workbook();
                AsposeCells.Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells.ImportData(table, 0, 0, new AsposeCells.ImportTableOptions { IsFieldNameShown = true });
                workbook.Save(outputStem + ".xlsx");
                break;
            }
            case "Export DataTable":
            {
                var workbook = new AsposeCells.Workbook(sourcePath);
                AsposeCells.Cells cells = workbook.Worksheets[0].Cells;
                DataTable exported = cells.ExportDataTableAsString(0, 0, cells.MaxDataRow + 1, cells.MaxDataColumn + 1, true);
                if (exported.Rows.Count < 1000)
                {
                    throw new InvalidOperationException($"Unexpected export row count {exported.Rows.Count}");
                }
                break;
            }
            case "Excel to CSV":
            {
                var workbook = new AsposeCells.Workbook(sourcePath);
                workbook.Save(outputStem + ".csv", AsposeCells.SaveFormat.Csv);
                break;
            }
            case "Excel to JSON":
            {
                var workbook = new AsposeCells.Workbook(sourcePath);
                AsposeCells.Cells cells = workbook.Worksheets[0].Cells;
                JsonExport.Write(outputStem + ".json", cells.MaxDataRow + 1, cells.MaxDataColumn + 1, (r, c) => cells[r - 1, c - 1].StringValue);
                break;
            }
        }
    }

    static void WriteGrid(AsposeCells.Worksheet sheet, DataTable table)
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

        sheet.Cells.ImportTwoDimensionArray(values, 0, 0);
    }
}

