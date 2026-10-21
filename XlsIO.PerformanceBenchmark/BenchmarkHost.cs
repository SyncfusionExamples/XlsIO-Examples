using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;

static class BenchmarkHost
{
    public const int FullRowCount = 100_000;
    public const int FreeRowCount = 148;
    public const int ColumnCount = 50;

    public static int Main(string[] args)
    {
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: XlsIO.PerformanceBenchmark <Library> <Method>");
            Console.WriteLine("Libraries: XlsIO, Aspose, DevExpress, GemBox, IronXL");
            Console.WriteLine("Methods: OpenAndSave, CreateAndSave, ImportDataTable, ExportDataTable, ExcelToCsv, ExcelToJson, ExcelToPdf");
            Console.WriteLine("GemBox and IronXL use the 150-row free-edition limit. The others use 100,000 rows.");
            return 1;
        }

        string library = args[0];
        string method = args[1];
        bool freeEdition = library.Equals("GemBox", StringComparison.OrdinalIgnoreCase)
            || library.Equals("IronXL", StringComparison.OrdinalIgnoreCase)
            || string.Equals(Environment.GetEnvironmentVariable("BENCH_FREE"), "1", StringComparison.Ordinal);
        int rowCount = freeEdition ? FreeRowCount : FullRowCount;

        string root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
        string dataDir = Path.Combine(root, "BenchmarkData");
        string resultsDir = Path.Combine(root, "BenchmarkResults");
        Directory.CreateDirectory(dataDir);
        Directory.CreateDirectory(resultsDir);

        RegisterFreeLicenses();
        DataTable table = BuildTable(rowCount, ColumnCount);
        string sourcePath = Path.Combine(dataDir, freeEdition ? "source-150.xlsx" : "source-100k.xlsx");
        if (!File.Exists(sourcePath))
        {
            WriteSourceWorkbook(sourcePath, table);
        }

        string outputStem = Path.Combine(dataDir, $"{Sanitize(library)}-{Sanitize(method)}");
        if (string.Equals(Environment.GetEnvironmentVariable("BENCH_WARMUP"), "1", StringComparison.Ordinal))
        {
            Invoke(library, method, sourcePath, table, outputStem + "-warmup");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        var process = Process.GetCurrentProcess();
        process.Refresh();
        long beforeWs = process.WorkingSet64;
        long beforeManaged = GC.GetTotalAllocatedBytes(precise: true);
        var clock = Stopwatch.StartNew();
        string status = "ok";
        string? error = null;
        long peakWs = beforeWs;

        try
        {
            Invoke(library, method, sourcePath, table, outputStem);
        }
        catch (Exception ex)
        {
            status = "error";
            error = ex.GetType().Name + ": " + ex.Message;
        }

        clock.Stop();
        process.Refresh();
        peakWs = Math.Max(peakWs, process.WorkingSet64);
        long allocated = GC.GetTotalAllocatedBytes(precise: false) - beforeManaged;

        var result = new BenchResult(
            library,
            method,
            rowCount,
            ColumnCount,
            status,
            clock.Elapsed.TotalMilliseconds,
            peakWs / (1024.0 * 1024.0),
            (peakWs - beforeWs) / (1024.0 * 1024.0),
            Math.Max(0, allocated) / (1024.0 * 1024.0),
            error);

        Console.WriteLine($"{result.Library} {result.Method} rows={result.Rows} {result.Status} time={result.ElapsedMs:F0} ms peakWS={result.PeakWorkingSetMb:F1} MB deltaWS={result.DeltaWorkingSetMb:F1} MB managed={result.AllocatedMb:F1} MB");
        if (error is not null)
        {
            Console.WriteLine(error);
        }

        string stamp = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
        string jsonPath = Path.Combine(resultsDir, $"run-{Sanitize(library)}-{Sanitize(method)}-{stamp}.json");
        File.WriteAllText(jsonPath, JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
        return status == "ok" ? 0 : 2;
    }

    static void Invoke(string library, string method, string sourcePath, DataTable table, string outputStem)
    {
        switch (library.ToUpperInvariant())
        {
            case "XLSIO":
                SyncfusionXlsIO.Run(method, sourcePath, table, outputStem);
                break;
            case "ASPOSE":
                AsposeCellsSample.Run(method, sourcePath, table, outputStem);
                break;
            case "DEVEXPRESS":
                DevExpressSpreadsheet.Run(method, sourcePath, table, outputStem);
                break;
            case "GEMBOX":
                GemBoxSpreadsheet.Run(method, sourcePath, table, outputStem);
                break;
            case "IRONXL":
                IronXlSample.Run(method, sourcePath, table, outputStem);
                break;
            default:
                throw new ArgumentException($"Unknown library '{library}'.");
        }
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

    static void RegisterFreeLicenses()
    {
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
    }

    static string Sanitize(string value) => value.Replace(" ", "", StringComparison.Ordinal);

    public static void EnsureRows(int actual, int expected)
    {
        if (actual < expected || actual > expected + 2)
        {
            throw new InvalidOperationException($"Unexpected row count {actual}; expected {expected}.");
        }
    }
}

sealed record BenchResult(
    string Library,
    string Method,
    int Rows,
    int Columns,
    string Status,
    double ElapsedMs,
    double PeakWorkingSetMb,
    double DeltaWorkingSetMb,
    double AllocatedMb,
    string? Error);
