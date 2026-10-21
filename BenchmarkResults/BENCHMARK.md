# Excel library performance comparison

- Runtime: .NET 9.0.21
- OS: Debian GNU/Linux 13 (trixie)
- Data: 100,000 rows x 50 columns (45.1 MB xlsx)
- Method: 0 warmup + 1 measured run.
- Time is elapsed milliseconds. Memory columns are megabytes.
- Delta WS is working-set growth during the timed region. Allocated is managed bytes allocated during the timed region.
- Competitors run under their public NuGet evaluation/trial terms. Failures caused by trial limits are shown as errors.

- Export DataTable was remeasured after fixing Aspose.Cells mixed-type export (`ExportDataTableAsString`). Other scenarios are from the same machine, single measured run, no warmup.
- Allocated MB of 0 means the allocator counter wrapped or did not advance; trust Delta WS in that case.
- Peak WS includes libraries already loaded in the process, so compare Delta WS for memory growth.

| Package | Version |
| --- | --- |
| Syncfusion.XlsIO.Net.Core | 35.1.41 |
| Aspose.Cells | 26.10.0 |
| DevExpress.Document.Processor | 26.1.5 |
| IronXL.Excel | 2026.10.5 |
| GemBox.Spreadsheet | 2026.10.116 |

## Open and Save

| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |
| --- | --- | ---: | ---: | ---: | ---: |
| Syncfusion XlsIO | ok | 17924 | 1040.0 | 869.8 | 3051.1 |
| Aspose.Cells | ok | 8773 | 967.0 | 730.1 | 920.8 |
| DevExpress Spreadsheet | ok | 17671 | 1322.1 | 1044.6 | 2394.0 |
| IronXL | error |  |  |  |  |
| GemBox.Spreadsheet | error |  |  |  |  |

Fastest: **Aspose.Cells** (8773 ms).

`IronXL` error: Production license required. Set `IRONXL_LICENSE_KEY` to a trial or production key. No timing was invented.

`GemBox.Spreadsheet` error: Free edition is limited to 150 rows. 100,000 rows require a professional license.

## Create and Save

| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |
| --- | --- | ---: | ---: | ---: | ---: |
| Syncfusion XlsIO | ok | 33028 | 1092.9 | 667.4 | 13486.3 |
| Aspose.Cells | ok | 8827 | 1440.9 | 1009.3 | 4014.4 |
| DevExpress Spreadsheet | ok | 20571 | 1336.5 | 889.6 | 4161.2 |
| IronXL | error |  |  |  |  |
| GemBox.Spreadsheet | error |  |  |  |  |

Fastest: **Aspose.Cells** (8827 ms).

`IronXL` error: Production license required. Set `IRONXL_LICENSE_KEY` to a trial or production key. No timing was invented.

`GemBox.Spreadsheet` error: You are trying to modify 100000 cells while ExcelFile.GroupMethodsAffectedCellsLimit is set to 5000. If you want to perform this operation, you need to change ExcelFile.GroupMethodsAffectedCellsLimit accordingly.

## Import DataTable

| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |
| --- | --- | ---: | ---: | ---: | ---: |
| Syncfusion XlsIO | ok | 12653 | 1012.9 | 554.4 | 7291.8 |
| Aspose.Cells | ok | 8577 | 1154.1 | 690.3 | 2295.9 |
| DevExpress Spreadsheet | ok | 20093 | 1264.5 | 800.3 | 4123.0 |
| IronXL | error |  |  |  |  |
| GemBox.Spreadsheet | error |  |  |  |  |

Fastest: **Aspose.Cells** (8577 ms).

`IronXL` error: Production license required. Set `IRONXL_LICENSE_KEY` to a trial or production key. No timing was invented.

`GemBox.Spreadsheet` error: Free edition is limited to 150 rows. 100,000 rows require a professional license.

## Export DataTable

| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |
| --- | --- | ---: | ---: | ---: | ---: |
| Syncfusion XlsIO | ok | 9961 | 1169.0 | 999.1 | 2341.4 |
| Aspose.Cells | ok | 12088 | 1361.1 | 1120.1 | 2018.8 |
| DevExpress Spreadsheet | ok | 7975 | 1169.9 | 893.4 | 1177.0 |
| IronXL | error |  |  |  |  |
| GemBox.Spreadsheet | error |  |  |  |  |

Fastest: **DevExpress Spreadsheet** (7975 ms).

`IronXL` error: Production license required. The public NuGet package refuses to run without `IRONXL_LICENSE_KEY` (trial or production). No timing was invented.

`GemBox.Spreadsheet` error: Free edition stops at 150 rows (`Unexpected export row count 149` after the free-limit handler). 100,000 rows require a professional license.

## Excel to CSV

| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |
| --- | --- | ---: | ---: | ---: | ---: |
| Syncfusion XlsIO | ok | 5452 | 1087.7 | 616.6 | 1818.0 |
| Aspose.Cells | ok | 4671 | 1167.5 | 696.0 | 1394.8 |
| DevExpress Spreadsheet | ok | 6727 | 1218.6 | 746.9 | 1145.2 |
| IronXL | error |  |  |  |  |
| GemBox.Spreadsheet | error |  |  |  |  |

Fastest: **Aspose.Cells** (4671 ms).

`IronXL` error: Production license required. Set `IRONXL_LICENSE_KEY` to a trial or production key. No timing was invented.

`GemBox.Spreadsheet` error: Free edition is limited to 150 rows. 100,000 rows require a professional license.

## Excel to JSON

| Library | Status | Median ms | Peak WS MB | Delta WS MB | Allocated MB |
| --- | --- | ---: | ---: | ---: | ---: |
| Syncfusion XlsIO | ok | 17751 | 1048.5 | 576.5 | 0.0 |
| Aspose.Cells | ok | 5325 | 1341.6 | 868.3 | 12171.7 |
| DevExpress Spreadsheet | ok | 12919 | 1527.8 | 1053.8 | 0.0 |
| IronXL | error |  |  |  |  |
| GemBox.Spreadsheet | error |  |  |  |  |

Fastest: **Aspose.Cells** (5325 ms).

`IronXL` error: Production license required. Set `IRONXL_LICENSE_KEY` to a trial or production key. No timing was invented.

`GemBox.Spreadsheet` error: Free edition is limited to 150 rows. 100,000 rows require a professional license.

