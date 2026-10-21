# Benchmark

Each library and method runs in its own process:

```text
dotnet run -- <Library> <Method>
```

Libraries: `XlsIO`, `Aspose`, `DevExpress`, `GemBox`, `IronXL`.

Methods: `OpenAndSave`, `CreateAndSave`, `ImportDataTable`, `ExportDataTable`, `ExcelToCsv`, `ExcelToJson`, `ExcelToPdf`.

XlsIO, Aspose, and DevExpress use 100,000 rows by 50 columns. GemBox and IronXL use 148 data rows so the worksheet stays inside the GemBox free edition limit of 150 rows, including the header. `BENCH_FREE=1` runs any library at that size. `BENCH_WARMUP=1` runs the method once before the timed call.
