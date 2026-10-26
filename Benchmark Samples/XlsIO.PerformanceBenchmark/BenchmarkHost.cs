using System.Data;
using System.Diagnostics;

static class BenchmarkHost
{
    public const int RowCount = 100_000;
    public const int ColumnCount = 50;

    public static int Main(string[] args)
    {
        if (args.Length < 1)
        {
            Console.WriteLine("Usage: XlsIO.PerformanceBenchmark <Method>");
            Console.WriteLine("Methods: OpenAndSave, CreateAndSave, ImportDataTable, ExportDataTable, ExcelToCsv, ExcelToJson, ExcelToPdf");
            return 1;
        }

        string method = args[0];
        string dataDir = Path.Combine(AppContext.BaseDirectory, "Output");
        Directory.CreateDirectory(dataDir);

        DataTable table = BuildTable(RowCount, ColumnCount);
        string sourcePath = Path.Combine(dataDir, "source.xlsx");
        if (!File.Exists(sourcePath))
        {
            WriteSourceWorkbook(sourcePath, table);
        }

        var clock = Stopwatch.StartNew();
        SyncfusionXlsIO.Run(method, sourcePath, table, Path.Combine(dataDir, method));
        clock.Stop();

        Console.WriteLine($"{method} rows={RowCount} time={clock.Elapsed.TotalMilliseconds:F0} ms");
        return 0;
    }

    public static DataTable BuildTable(int rows, int columns)
    {
        var table = new DataTable("Sales");
        for (int c = 0; c < columns; c++)
        {
            table.Columns.Add(c % 5 == 4 ? $"Amount{c}" : $"Col{c}", c % 5 == 4 ? typeof(double) : typeof(string));
        }

        for (int r = 0; r < rows; r++)
        {
            var values = new object[columns];
            for (int c = 0; c < columns; c++)
            {
                values[c] = c % 5 == 4 ? (object)((r * 17 + c) % 10000 / 100.0) : $"R{r}C{c}";
            }

            table.Rows.Add(values);
        }

        return table;
    }

    static void WriteSourceWorkbook(string path, DataTable table)
    {
        using var engine = new Syncfusion.XlsIO.ExcelEngine();
        Syncfusion.XlsIO.IApplication application = engine.Excel;
        application.DefaultVersion = Syncfusion.XlsIO.ExcelVersion.Xlsx;
        Syncfusion.XlsIO.IWorkbook workbook = application.Workbooks.Create(1);
        workbook.Worksheets[0].ImportDataTable(table, true, 1, 1);
        using FileStream stream = File.Create(path);
        workbook.SaveAs(stream);
        workbook.Close();
    }

    public static void EnsureRows(int actual, int expected)
    {
        if (actual < expected || actual > expected + 2)
        {
            throw new InvalidOperationException($"Unexpected row count {actual}; expected {expected}.");
        }
    }
}
