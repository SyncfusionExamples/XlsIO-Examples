# XlsIO-Examples
This repository contains examples of creating, reading, editing, and converting Excel documents programmatically using Syncfusion XlsIO.

## Performance benchmark

`XlsIO.PerformanceBenchmark` is a .NET 9 console app that measures elapsed time and memory for 100,000 rows by 50 columns:

- Open and Save
- Create and Save
- Import DataTable
- Export DataTable
- Excel to CSV
- Excel to JSON

Libraries: Syncfusion XlsIO, Aspose.Cells, DevExpress Spreadsheet, IronXL, and GemBox.Spreadsheet (latest NuGet versions).

```bash
dotnet run --project XlsIO.PerformanceBenchmark -c Release
```

Optional: `BENCH_ONLY="Open and Save,Excel to CSV"` runs a subset. `IRONXL_LICENSE_KEY` supplies an IronXL trial or production key. Results are written to `BenchmarkResults/`.

The measured comparison from this environment is in [BenchmarkResults/BENCHMARK.md](BenchmarkResults/BENCHMARK.md). IronXL and GemBox.Spreadsheet cannot process 100,000 rows under their public NuGet trial limits, so those rows are recorded as errors rather than timings.
