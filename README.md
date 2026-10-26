# XlsIO-Examples
This repository contains examples of creating, reading, editing, and converting Excel documents programmatically using Syncfusion XlsIO.

## Performance benchmark

`XlsIO.PerformanceBenchmark` is a .NET 9 console app that measures the elapsed time of Syncfusion XlsIO operations on a workbook with 100,000 rows by 50 columns. Each run executes one method:

- `OpenAndSave`
- `CreateAndSave`
- `ImportDataTable`
- `ExportDataTable`
- `ExcelToCsv`
- `ExcelToJson`
- `ExcelToPdf`

```bash
dotnet run --project XlsIO.PerformanceBenchmark -c Release -- ExcelToCsv
```

Generated files are written to the `Output` folder next to the built application.

NuGet packages: `Syncfusion.XlsIO.Net.Core`, `Syncfusion.XlsIORenderer.Net.Core`, and `Syncfusion.Pdf.Net.Core` (35.1.41).
