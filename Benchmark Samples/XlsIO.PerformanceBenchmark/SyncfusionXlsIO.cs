using System.Data;
using Syncfusion.XlsIO;
using Syncfusion.XlsIORenderer;
using Syncfusion.Pdf;

static class SyncfusionXlsIO
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
        using var engine = new ExcelEngine();
        IApplication application = engine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;
        IWorkbook workbook = application.Workbooks.Open(sourcePath);
        using FileStream stream = File.Create(outputStem + ".xlsx");
        workbook.SaveAs(stream);
        workbook.Close();
    }

    public static void CreateAndSave(DataTable table, string outputStem)
    {
        using var engine = new ExcelEngine();
        IApplication application = engine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;
        IWorkbook workbook = application.Workbooks.Create(1);
        IWorksheet sheet = workbook.Worksheets[0];
        WriteGrid(sheet, table);
        using FileStream stream = File.Create(outputStem + ".xlsx");
        workbook.SaveAs(stream);
        workbook.Close();
    }

    public static void ImportDataTable(DataTable table, string outputStem)
    {
        using var engine = new ExcelEngine();
        IApplication application = engine.Excel;
        application.DefaultVersion = ExcelVersion.Xlsx;
        IWorkbook workbook = application.Workbooks.Create(1);
        workbook.Worksheets[0].ImportDataTable(table, true, 1, 1);
        using FileStream stream = File.Create(outputStem + ".xlsx");
        workbook.SaveAs(stream);
        workbook.Close();
    }

    public static void ExportDataTable(string sourcePath, int expectedRows)
    {
        using var engine = new ExcelEngine();
        IApplication application = engine.Excel;
        IWorkbook workbook = application.Workbooks.Open(sourcePath);
        DataTable exported = workbook.Worksheets[0].ExportDataTable(workbook.Worksheets[0].UsedRange, ExcelExportDataTableOptions.ColumnNames);
        BenchmarkHost.EnsureRows(exported.Rows.Count, expectedRows);
        workbook.Close();
    }

    public static void ExcelToCsv(string sourcePath, string outputStem)
    {
        using var engine = new ExcelEngine();
        IApplication application = engine.Excel;
        IWorkbook workbook = application.Workbooks.Open(sourcePath);
        workbook.SaveAs(outputStem + ".csv", ",");
        workbook.Close();
    }

    public static void ExcelToJson(string sourcePath, string outputStem)
    {
        using var engine = new ExcelEngine();
        IApplication application = engine.Excel;
        IWorkbook workbook = application.Workbooks.Open(sourcePath);
        workbook.SaveAsJson(outputStem + ".json", workbook.Worksheets[0]);
        workbook.Close();
    }

    public static void ExcelToPdf(string sourcePath, string outputStem)
    {
        using var engine = new ExcelEngine();
        IApplication application = engine.Excel;
        IWorkbook workbook = application.Workbooks.Open(sourcePath);
        var renderer = new XlsIORenderer();
        using PdfDocument pdf = renderer.ConvertToPDF(workbook);
        using FileStream stream = File.Create(outputStem + ".pdf");
        pdf.Save(stream);
        workbook.Close();
    }

    static void WriteGrid(IWorksheet sheet, DataTable table)
    {
        int columns = table.Columns.Count;
        for (int c = 0; c < columns; c++)
        {
            sheet.SetValue(1, c + 1, table.Columns[c].ColumnName);
        }

        for (int r = 0; r < table.Rows.Count; r++)
        {
            int row = r + 2;
            DataRow dataRow = table.Rows[r];
            for (int c = 0; c < columns; c++)
            {
                sheet.SetValue(row, c + 1, dataRow[c] is double amount ? amount.ToString(System.Globalization.CultureInfo.InvariantCulture) : (string)dataRow[c]);
            }
        }
    }
}
