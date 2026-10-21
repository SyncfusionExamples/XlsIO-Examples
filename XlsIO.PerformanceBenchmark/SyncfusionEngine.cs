using System.Data;
using Syncfusion.XlsIO;

sealed class SyncfusionEngine : IExcelEngine
{
    public string Name => "Syncfusion XlsIO";

    public void Run(string scenario, string sourcePath, DataTable table, string outputStem)
    {
        using ExcelEngine excelEngine = new ExcelEngine();
        IApplication application = excelEngine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;

        switch (scenario)
        {
            case "Open and Save":
            {
                using FileStream input = File.OpenRead(sourcePath);
                IWorkbook workbook = application.Workbooks.Open(input);
                using FileStream output = File.Create(outputStem + ".xlsx");
                workbook.SaveAs(output);
                workbook.Close();
                break;
            }
            case "Create and Save":
            {
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];
                WriteGrid(sheet, table);
                using FileStream output = File.Create(outputStem + ".xlsx");
                workbook.SaveAs(output);
                workbook.Close();
                break;
            }
            case "Import DataTable":
            {
                IWorkbook workbook = application.Workbooks.Create(1);
                workbook.Worksheets[0].ImportDataTable(table, true, 1, 1);
                using FileStream output = File.Create(outputStem + ".xlsx");
                workbook.SaveAs(output);
                workbook.Close();
                break;
            }
            case "Export DataTable":
            {
                using FileStream input = File.OpenRead(sourcePath);
                IWorkbook workbook = application.Workbooks.Open(input);
                IWorksheet sheet = workbook.Worksheets[0];
                DataTable exported = sheet.ExportDataTable(sheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);
                if (exported.Rows.Count < 1000)
                {
                    throw new InvalidOperationException($"Unexpected export row count {exported.Rows.Count}");
                }
                workbook.Close();
                break;
            }
            case "Excel to CSV":
            {
                using FileStream input = File.OpenRead(sourcePath);
                IWorkbook workbook = application.Workbooks.Open(input);
                workbook.SaveAs(outputStem + ".csv", ",");
                workbook.Close();
                break;
            }
            case "Excel to JSON":
            {
                using FileStream input = File.OpenRead(sourcePath);
                IWorkbook workbook = application.Workbooks.Open(input);
                IWorksheet sheet = workbook.Worksheets[0];
                JsonExport.Write(outputStem + ".json", sheet.UsedRange.LastRow, sheet.UsedRange.LastColumn, (r, c) => sheet.Range[r, c].DisplayText);
                workbook.Close();
                break;
            }
        }
    }

    static void WriteGrid(IWorksheet sheet, DataTable table)
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

        sheet.ImportArray(values, 1, 1);
    }
}

