using Syncfusion.XlsIO;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json;

const int RowCount = 100_000;
const int ColumnCount = 50;
const int WarmupIterations = 0;
const int MeasuredIterations = 1;

string root = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
string dataDir = Path.Combine(root, "BenchmarkData");
string resultsDir = Path.Combine(root, "BenchmarkResults");
Directory.CreateDirectory(dataDir);
Directory.CreateDirectory(resultsDir);

Console.WriteLine($"Rows={RowCount:N0} Columns={ColumnCount:N0}");
Console.WriteLine($"Process working set at start: {Process.GetCurrentProcess().WorkingSet64 / (1024 * 1024)} MB");
Console.WriteLine("Building shared DataTable and source workbook...");

DataTable table = BuildTable(RowCount, ColumnCount);
string sourcePath = Path.Combine(dataDir, "source-100k.xlsx");
if (!File.Exists(sourcePath))
{
    WriteSourceWorkbook(sourcePath, table);
}
long sourceBytes = new FileInfo(sourcePath).Length;
Console.WriteLine($"Source workbook: {sourcePath} ({sourceBytes / (1024.0 * 1024.0):F1} MB)");

RegisterFreeLicenses();

var engines = new IExcelEngine[]
{
    new SyncfusionEngine(),
    new AsposeEngine(),
    new DevExpressEngine(),
    new IronXlEngine(),
    new GemBoxEngine(),
};

string? engineFilter = Environment.GetEnvironmentVariable("BENCH_ENGINE");
if (!string.IsNullOrWhiteSpace(engineFilter))
{
    engines = engines.Where(engine => engine.Name.Contains(engineFilter, StringComparison.OrdinalIgnoreCase)).ToArray();
}

string[] scenarios = string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("BENCH_ONLY"))
    ? [
        "Open and Save",
        "Create and Save",
        "Import DataTable",
        "Export DataTable",
        "Excel to CSV",
        "Excel to JSON",
    ]
    : Environment.GetEnvironmentVariable("BENCH_ONLY")!.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

var results = new List<BenchResult>();

foreach (string scenario in scenarios)
{
    Console.WriteLine();
    Console.WriteLine($"======== {scenario} ========");
    foreach (IExcelEngine engine in engines)
    {
        BenchResult result = Measure(engine, scenario, sourcePath, table, dataDir);
        results.Add(result);
        Console.WriteLine(
            $"{engine.Name,-28} {result.Status,-8} time={FormatMs(result.MedianMs),10}  peakWS={FormatMb(result.PeakWorkingSetMb),10}  deltaWS={FormatMb(result.DeltaWorkingSetMb),10}  managed={FormatMb(result.AllocatedMb),10}");
        if (result.Error is not null)
        {
            Console.WriteLine($"    {result.Error}");
        }
    }
}

string stamp = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture);
string jsonPath = Path.Combine(resultsDir, $"benchmark-{stamp}.json");
string mdPath = Path.Combine(resultsDir, $"benchmark-{stamp}.md");
string latestMd = Path.Combine(resultsDir, "BENCHMARK.md");

var payload = new
{
    generatedUtc = DateTime.UtcNow,
    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
    processorCount = Environment.ProcessorCount,
    rows = RowCount,
    columns = ColumnCount,
    warmupIterations = WarmupIterations,
    measuredIterations = MeasuredIterations,
    sourceWorkbookBytes = sourceBytes,
    packages = new Dictionary<string, string>
    {
        ["Syncfusion.XlsIO.Net.Core"] = "35.1.41",
        ["Aspose.Cells"] = "26.10.0",
        ["DevExpress.Document.Processor"] = "26.1.5",
        ["IronXL.Excel"] = "2026.10.5",
        ["GemBox.Spreadsheet"] = "2026.10.116",
    },
    notes = new[]
    {
        "Median of measured iterations after one warmup. Each iteration is a fresh process-level GC collect before the timed region.",
        "PeakWorkingSetMb is the process peak working set observed during the timed region (includes other loaded libraries).",
        "DeltaWorkingSetMb is peak working set minus working set sampled immediately before the timed region.",
        "AllocatedMb is GC.GetTotalAllocatedBytes delta during the timed region (managed allocations, not peak heap).",
        "Aspose.Cells, DevExpress, IronXL, and GemBox run in evaluation/trial mode. Trial restrictions (watermarks, row limits, or disabled features) are reported as errors rather than fabricated timings.",
        "Excel to JSON writes a JSON array of row objects using each library's cell reader. It is not a vendor-specific JSON converter.",
    },
    results,
};

File.WriteAllText(jsonPath, JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true }));
string markdown = BuildMarkdown(payload.runtime, payload.os, sourceBytes, results, scenarios, engines.Select(e => e.Name).ToArray());
File.WriteAllText(mdPath, markdown);
File.WriteAllText(latestMd, markdown);
Console.WriteLine();
Console.WriteLine(markdown);
Console.WriteLine($"Wrote {jsonPath}");
Console.WriteLine($"Wrote {mdPath}");

static BenchResult Measure(IExcelEngine engine, string scenario, string sourcePath, DataTable table, string dataDir)
{
    string outDir = Path.Combine(dataDir, Sanitize(engine.Name), Sanitize(scenario));
    Directory.CreateDirectory(outDir);

    try
    {
        for (int i = 0; i < WarmupIterations; i++)
        {
            engine.Run(scenario, sourcePath, table, Path.Combine(outDir, $"warmup-{i}"));
            Cleanup();
        }

        var samples = new List<(double ms, double peak, double delta, double allocated)>();
        for (int i = 0; i < MeasuredIterations; i++)
        {
            Cleanup();
            long beforeWs = Process.GetCurrentProcess().WorkingSet64;
            long beforeAlloc = GC.GetTotalAllocatedBytes(precise: false);
            long beforePeak = ProcessPeakWorkingSet();

            var sw = Stopwatch.StartNew();
            engine.Run(scenario, sourcePath, table, Path.Combine(outDir, $"run-{i}"));
            sw.Stop();

            long afterPeak = ProcessPeakWorkingSet();
            long afterWs = Process.GetCurrentProcess().WorkingSet64;
            long afterAlloc = GC.GetTotalAllocatedBytes(precise: false);
            double peakMb = Math.Max(afterPeak, afterWs) / (1024.0 * 1024.0);
            double deltaMb = Math.Max(0, Math.Max(afterPeak, afterWs) - beforeWs) / (1024.0 * 1024.0);
            // Peak counters are process-lifetime on some OSes; fall back to end working set delta.
            if (afterPeak <= beforePeak)
            {
                peakMb = afterWs / (1024.0 * 1024.0);
                deltaMb = Math.Max(0, afterWs - beforeWs) / (1024.0 * 1024.0);
            }

            samples.Add((sw.Elapsed.TotalMilliseconds, peakMb, deltaMb, (afterAlloc - beforeAlloc) / (1024.0 * 1024.0)));
            Cleanup();
        }

        samples.Sort((a, b) => a.ms.CompareTo(b.ms));
        var median = samples[samples.Count / 2];
        return new BenchResult(engine.Name, scenario, "ok", median.ms, median.peak, median.delta, median.allocated, null);
    }
    catch (Exception ex)
    {
        return new BenchResult(engine.Name, scenario, "error", 0, 0, 0, 0, Flatten(ex));
    }
}

static void Cleanup()
{
    GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
    GC.WaitForPendingFinalizers();
    GC.Collect(GC.MaxGeneration, GCCollectionMode.Aggressive, blocking: true, compacting: true);
}

static long ProcessPeakWorkingSet()
{
    try
    {
        return Process.GetCurrentProcess().PeakWorkingSet64;
    }
    catch
    {
        return Process.GetCurrentProcess().WorkingSet64;
    }
}

static DataTable BuildTable(int rows, int columns)
{
    var table = new DataTable("Sales");
    for (int c = 0; c < columns; c++)
    {
        Type type = (c % 5) switch
        {
            0 => typeof(int),
            1 => typeof(double),
            2 => typeof(string),
            3 => typeof(DateTime),
            _ => typeof(decimal),
        };
        table.Columns.Add($"C{c:00}", type);
    }

    var random = new Random(42);
    for (int r = 0; r < rows; r++)
    {
        var values = new object[columns];
        for (int c = 0; c < columns; c++)
        {
            values[c] = (c % 5) switch
            {
                0 => r + c,
                1 => Math.Round(random.NextDouble() * 1000, 2),
                2 => $"R{r}-C{c}",
                3 => new DateTime(2020, 1, 1).AddDays(r % 2000),
                _ => Math.Round((decimal)random.NextDouble() * 500m, 2),
            };
        }
        table.Rows.Add(values);
    }

    return table;
}

static void WriteSourceWorkbook(string path, DataTable table)
{
    using ExcelEngine excelEngine = new ExcelEngine();
    IApplication application = excelEngine.Excel;
    application.DefaultVersion = ExcelVersion.Xlsx;
    IWorkbook workbook = application.Workbooks.Create(1);
    IWorksheet worksheet = workbook.Worksheets[0];
    worksheet.Name = "Sales";
    worksheet.ImportDataTable(table, true, 1, 1);
    using FileStream stream = File.Create(path);
    workbook.SaveAs(stream);
    workbook.Close();
}

static void RegisterFreeLicenses()
{
    try
    {
        GemBox.Spreadsheet.SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        GemBox.Spreadsheet.SpreadsheetInfo.FreeLimitReached += (_, args) => args.FreeLimitReachedAction = GemBox.Spreadsheet.FreeLimitReachedAction.Stop;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"GemBox license note: {ex.Message}");
    }
}

static string BuildMarkdown(string runtime, string os, long sourceBytes, List<BenchResult> results, string[] scenarios, string[] engines)
{
    var sb = new StringBuilder();
    sb.AppendLine("# Excel library performance comparison");
    sb.AppendLine();
    sb.AppendLine($"- Runtime: {runtime}");
    sb.AppendLine($"- OS: {os}");
    sb.AppendLine($"- Data: {RowCount:N0} rows x {ColumnCount} columns ({sourceBytes / (1024.0 * 1024.0):F1} MB xlsx)");
    sb.AppendLine($"- Method: {WarmupIterations} warmup + {MeasuredIterations} measured run.");
    sb.AppendLine("- Time is elapsed milliseconds. Memory columns are megabytes.");
    sb.AppendLine("- Delta WS is working-set growth during the timed region. Allocated is managed bytes allocated during the timed region.");
    sb.AppendLine("- Competitors run under their public NuGet evaluation/trial terms. Failures caused by trial limits are shown as errors.");
    sb.AppendLine();
    sb.AppendLine("| Package | Version |");
    sb.AppendLine("| --- | --- |");
    sb.AppendLine("| Syncfusion.XlsIO.Net.Core | 35.1.41 |");
    sb.AppendLine("| Aspose.Cells | 26.10.0 |");
    sb.AppendLine("| DevExpress.Document.Processor | 26.1.5 |");
    sb.AppendLine("| IronXL.Excel | 2026.10.5 |");
    sb.AppendLine("| GemBox.Spreadsheet | 2026.10.116 |");
    sb.AppendLine();

    foreach (string scenario in scenarios)
    {
        sb.AppendLine($"## {scenario}");
        sb.AppendLine();
        sb.AppendLine("| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |");
        sb.AppendLine("| --- | --- | ---: | ---: | ---: | ---: |");
        foreach (string engine in engines)
        {
            BenchResult result = results.Single(r => r.Library == engine && r.Scenario == scenario);
            if (result.Status != "ok")
            {
                sb.AppendLine($"| {engine} | error |  |  |  |  |");
                continue;
            }

            sb.AppendLine($"| {engine} | ok | {result.MedianMs:F0} | {result.PeakWorkingSetMb:F1} | {result.DeltaWorkingSetMb:F1} | {result.AllocatedMb:F1} |");
        }

        var ok = results.Where(r => r.Scenario == scenario && r.Status == "ok").OrderBy(r => r.MedianMs).ToList();
        if (ok.Count > 0)
        {
            sb.AppendLine();
            sb.AppendLine($"Fastest: **{ok[0].Library}** ({ok[0].MedianMs:F0} ms).");
        }

        foreach (BenchResult failed in results.Where(r => r.Scenario == scenario && r.Status != "ok"))
        {
            sb.AppendLine();
            sb.AppendLine($"`{failed.Library}` error: {failed.Error}");
        }

        sb.AppendLine();
    }

    return sb.ToString();
}

static string FormatMs(double value) => value <= 0 ? "-" : value.ToString("F0", CultureInfo.InvariantCulture);
static string FormatMb(double value) => value <= 0 ? "-" : value.ToString("F1", CultureInfo.InvariantCulture);
static string Sanitize(string value) => string.Concat(value.Select(ch => char.IsLetterOrDigit(ch) ? ch : '-'));
static string Flatten(Exception ex) => ex.GetBaseException().Message.Replace('\r', ' ').Replace('\n', ' ');

