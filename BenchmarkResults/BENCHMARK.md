# Excel library performance comparison

- Runtime: .NET 9.0.21
- OS: Debian GNU/Linux 13 (trixie)
- Data: 100,000 rows x 50 columns (45.1 MB xlsx)
- Method: 0 warmup + 1 measured run.
- Time is elapsed milliseconds. Memory columns are megabytes.
- Delta WS is working-set growth during the timed region. Allocated is managed bytes allocated during the timed region.
- Competitors run under their public NuGet evaluation/trial terms. Failures caused by trial limits are shown as errors.

| Package | Version |
| --- | --- |
| Syncfusion.XlsIO.Net.Core | 35.1.41 |
| Aspose.Cells | 26.10.0 |
| DevExpress.Document.Processor | 26.1.5 |
| IronXL.Excel | 2026.10.5 |
| GemBox.Spreadsheet | 2026.10.116 |

## Excel to JSON

| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |
| --- | --- | ---: | ---: | ---: | ---: |
| Syncfusion XlsIO | ok | 20136 | 935.5 | 765.0 | 14513.0 |

Fastest: **Syncfusion XlsIO** (20136 ms).

