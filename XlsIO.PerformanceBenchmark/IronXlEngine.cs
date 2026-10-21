using System.Data;
using IronXL;

sealed class IronXlEngine : IExcelEngine
{
    public string Name => "IronXL";

    public void Run(string scenario, string sourcePath, DataTable table, string outputStem)
    {
        switch (scenario)
        {
            case "Open and Save":
            {
                WorkBook workbook = WorkBook.Load(sourcePath);
                workbook.SaveAs(outputStem + ".xlsx");
                break;
            }
            case "Create and Save":
            {
                WorkBook workbook = WorkBook.Create();
                WriteGrid(workbook.DefaultWorkSheet, table);
                workbook.SaveAs(outputStem + ".xlsx");
                break;
            }
            case "Import DataTable":
            {
                WorkBook workbook = WorkBook.Create();
                workbook.LoadWorkSheet(table);
                workbook.SaveAs(outputStem + ".xlsx");
                break;
            }
            case "Export DataTable":
            {
                WorkBook workbook = WorkBook.Load(sourcePath);
                DataTable exported = workbook.DefaultWorkSheet.ToDataTable(true);
                if (exported.Rows.Count < 1000)
                {
                    throw new InvalidOperationException($"Unexpected export row count {exported.Rows.Count}");
                }
                break;
            }
            case "Excel to CSV":
            {
                WorkBook workbook = WorkBook.Load(sourcePath);
                workbook.SaveAsCsv(outputStem + ".csv");
                break;
            }
            case "Excel to JSON":
            {
                WorkBook workbook = WorkBook.Load(sourcePath);
                WorkSheet sheet = workbook.DefaultWorkSheet;
                JsonExport.Write(outputStem + ".json", sheet.RowCount, sheet.ColumnCount, (r, c) => sheet.GetCellAt(r - 1, c - 1)?.StringValue ?? "");
                break;
            }
        }
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

