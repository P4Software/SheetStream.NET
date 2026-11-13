# SheetStream.NET

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![NuGet](https://img.shields.io/nuget/v/SheetStream.NET.svg)](https://www.nuget.org/packages/SheetStream.NET) [![Downloads](https://img.shields.io/nuget/dt/SheetStream.NET.svg)](https://www.nuget.org/packages/SheetStream.NET) [![Build](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml/badge.svg)](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml) [![Coverage](https://img.shields.io/endpoint?url=https%3A%2F%2Fp4software.github.io%2FSheetStream.NET%2Fcoverage.json)](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml) [![OpenSSF Scorecard](https://img.shields.io/ossf-scorecard/github.com/P4Software/SheetStream.NET?label=OpenSSF%20Scorecard)](https://scorecard.dev/viewer/?uri=github.com/P4Software/SheetStream.NET) [![Docs](https://img.shields.io/badge/docs-p4software.github.io-0A5FAD)](https://p4software.github.io/SheetStream.NET/)

SheetStream.NET is a zero-dependency, MIT-licensed .NET library that creates, fills, edits and reads Excel (`.xlsx`) files. It is not related to the Node.js package of the same name.

P4 Software built it when EPPlus moved to a commercial licence. Rows are written straight into the zip archive through an `XmlWriter`, so memory stays flat however many rows you export. It targets `netstandard2.0` and `net8.0`, which covers .NET Framework 4.7.2 and later, and .NET 8 and later. The namespace and assembly are `SheetStream`.

## Install

```bash
dotnet add package SheetStream.NET
dotnet add package SheetStream.AspNetCore   # optional: Excel downloads from ASP.NET Core
```

Coming from the 0.3.x `SheetStream` package: run `dotnet remove package SheetStream` and add `SheetStream.NET`. The namespace did not change. Documentation and API reference: https://p4software.github.io/SheetStream.NET/

## Quick start

```csharp
using SheetStream;
using SheetStream.Styles;

using (var package = ExcelPackage.Create("FinancialReport.xlsx"))
{
    var workbook = package.Workbook;
    var sheet = workbook.Worksheets.Add("Q1 Report");

    sheet.Column(1).SetWidth(10);
    sheet.Column(2).SetWidth(28);
    sheet.Column(3).SetWidth(18);
    sheet.FreezePanes(rows: 1);

    var headerStyle = workbook.Styles.Create(o =>
    {
        o.Font.Bold = true;
        o.Font.Color = "#FFFFFF";
        o.Fill.BackgroundColor = "#4F81BD";
        o.Border.Style = BorderStyle.Thin;
    });
    var currencyStyle = workbook.Styles.Create(o => o.NumberFormat = "$#,##0.00");

    sheet.Row(1)
        .WriteCell("ID", headerStyle)
        .WriteCell("Product Name", headerStyle)
        .WriteCell("Revenue (USD)", headerStyle);

    for (int i = 1; i <= 100; i++)
    {
        sheet.Row(i + 1)
            .WriteCell(i)
            .WriteCell($"Enterprise SKU-{i}")          // Shared string, de-duplicated
            .WriteCell(1500.50 * i, currencyStyle);
    }

    sheet.Row(102).WriteCell("Total", headerStyle).WriteEmpty(headerStyle).WriteFormula("SUM(C2:C101)", currencyStyle);
    package.Save();
}
```

To get bytes for a web response, or to load a whole collection or query in one call:

```csharp
byte[] bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Export").LoadFromCollection(orders));
byte[] fromQuery = await ExcelPackage.CreateBytesAsync(async wb =>
    wb.Worksheets.Add("Export").LoadFromCollection(await db.Orders.ToListAsync()));
```

## What it can do

The API reference on the docs site lists every member. These are the entry points.

| Area | Entry point | More |
|---|---|---|
| Load data | `LoadFromCollection`, `LoadFromDataTable`, `LoadFromDataReader`, `LoadFromArrays`, `LoadFromAsyncEnumerable` (.NET 8+) | [docs site](https://p4software.github.io/SheetStream.NET/) |
| Styles, tables, conditional formatting, validation | `Styles.Create`, `AddTable`, `AddConditionalFormatting`, `AddDataValidation` | [docs site](https://p4software.github.io/SheetStream.NET/) |
| Charts and pictures | `AddChart`, `AddPicture` | [docs site](https://p4software.github.io/SheetStream.NET/) |
| Pivot tables | `AddPivotTable` | [pivot-tables.md](docs/pivot-tables.md) |
| Links, comments, rich text, formulas | `WriteHyperlink`, `AddComment`, `WriteRichText`, `WriteFormula`, `CacheFormulaResults` | [docs index](docs/index.md) |
| Layout, printing, protection | `AutoFitColumns`, `FreezePanes`, `PageSetup`, `Protect` | [docs site](https://p4software.github.io/SheetStream.NET/) |
| Fill a template designed in Excel | `ExcelPackage.OpenTemplate` | [templates.md](docs/templates.md) |
| Edit a small existing workbook | `ExcelPackage.Edit` | [editing-workbooks.md](docs/editing-workbooks.md) |
| Trimming and Native AOT | `[SheetRow]` source generator | [aot-and-source-generator.md](docs/aot-and-source-generator.md) |
| Build-time warnings for streaming mistakes | Roslyn analyzers SS0001 to SS0003 | [analyzers.md](docs/analyzers.md) |
| Download from ASP.NET Core | `ExcelResults.Excel(fileName, build)` | [SheetStream.AspNetCore](src/SheetStream.AspNetCore/README.md) |

Formula cells show blank in file previews and in `ExcelReader` unless the file stores a result. `CacheFormulaResults()` computes simple formulas (`+ - * / %`, `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `ROUND`, `ABS` over cells already written on the sheet) and writes nothing for anything else.

Editing a file in place commits only on an explicit `Save()`:

```csharp
using var package = ExcelPackage.Edit("budget.xlsx");           // or Edit(input, output)
var sheet = package.Workbook.Worksheets.Edit("Summary");
sheet.Cells["B4"].Value = 1400;                                  // keeps the cell's number format
sheet.Cells["C4"].Formula = "B4*1.21";
package.Save();                                                  // atomic replace
```

Editing holds the edited sheet in memory and refuses sheets above `PackageOptions.MaxEditableCells` (200,000 cells by default).

## Reading

```csharp
using SheetStream.Reading;

using var reader = ExcelReader.Open("import.xlsx");
foreach (var row in reader.ReadRows("Products"))                      // streams, constant memory
    Console.WriteLine($"{row.GetString(0)}: {row.GetDouble(2)}");

List<Product> products = reader.Read<Product>("Products").ToList();   // header row maps to properties
DataTable table = reader.ReadDataTable("Products");
await foreach (var row in reader.ReadRowsAsync("Products", ct)) { }   // .NET 8+, also OpenAsync and ReadAsync<T>
```

Dates come back as `DateTime`, elapsed times as `TimeSpan`, errors as `ExcelErrorValue`, and formula cells as their cached result. `ReadMergedRanges` and `ReadHyperlinks` read only the end of the sheet part, so they are fast on large sheets.

### Files from people you do not know

An `.xlsx` is a zip archive, so a small upload can unpack into gigabytes. `ExcelReader` rejects such files with an `InvalidDataException` that names the limit. The defaults sit far above anything Excel writes; change them with `ReaderOptions`:

```csharp
using var reader = ExcelReader.Open(upload, leaveOpen: false, new ReaderOptions { MaxCellTextLength = 32_767 }); // Excel's own cell limit
```

A stream that cannot seek is copied first: up to 64 MB in memory, beyond that to a temporary file that is deleted with the reader. `OpenTemplate` and `Edit` load the whole source workbook and do not apply `ReaderOptions`; use them on files your application owns. See [SECURITY.md](SECURITY.md).

## Streaming rules

SheetStream.NET writes while you go, which is why memory stays flat. The cost is a few forward-only rules, each enforced with a clear exception (and, where the analyzer can prove it, a build warning):

1. One sheet at a time. Adding a worksheet seals the previous one.
2. Rows ascend. `Row(n)` must be the current row or a later one; gaps are fine.
3. Layout first. Column widths, frozen panes, tab colour, gridlines and fit-to-page are stored before the data, so set them before the first row. `AutoFitColumns()` and `Cells[...]` defer the header and lift this rule for that sheet.
4. A sheet uses `Row()` or `Cells[r, c]`, not both. `Cells[r, c]` gives EPPlus-style random access and buffers that one sheet in memory, so use it for small, irregular layouts.
5. Dispose saves, except an in-place edit. Leaving a `using` block without `Save()` still finishes a file you are creating. `ExcelPackage.Edit(path)` replaces an existing file, so it commits only on an explicit `Save()`; if your code throws first, the original stays untouched.

## Benchmarks

Writing 100,000 rows of 10 mixed, formatted columns. BenchmarkDotNet on .NET 8.0.31, Intel Core i7-8700, Windows 11, measured on one developer workstation (2026-10-13). Peak memory is the peak working set above an empty baseline process, from separate fresh-process runs.

| Library | Time | Allocated | Peak above baseline | File size |
|---|--:|--:|--:|--:|
| SheetStream.NET (default compression) | 1,012.6 ms | 26.67 MB | 29.6 MB | 5,798 KB |
| SheetStream.NET (`CompressionLevel.Fastest`) | 534.1 ms | 26.67 MB | 31.0 MB | 7,006 KB |
| MiniExcel 1.46.0 | 596.4 ms | 343.36 MB | 37.3 MB | 7,421 KB |
| NPOI 2.7.6 SXSSF | 2,039.7 ms | 704.22 MB | 65.2 MB | 5,518 KB |
| EPPlus 4.5.3.3 | 2,286.7 ms | 390.19 MB | 183.8 MB | 5,601 KB |
| ClosedXML 0.105.1 | 3,103.2 ms | 1,104.16 MB | 231.1 MB | 6,030 KB |
| NPOI 2.7.6 XSSF | 4,658.4 ms | 1,186.16 MB | 545.1 MB | 5,304 KB |

- At default settings MiniExcel is faster. SheetStream.NET defaults to the smallest files (`CompressionLevel.Optimal`); with `Fastest` the two are level on this sheet (534 ms against 596 ms is inside the error bars).
- Allocation is where SheetStream.NET is clearly ahead: about 13 times less than the next library, and the same at both compression levels.
- NPOI and EPPlus write slightly smaller files on this sheet.
- Reading the same 100,000-row file: SheetStream.NET 532.7 ms and 118.62 MB allocated, ExcelDataReader 915.9 ms and 211.03 MB, EPPlus 4.5.3.3 1,205.8 ms and 787.26 MB, MiniExcel 1,829.3 ms and 2,369.07 MB, ClosedXML 3,034.0 ms and 1,598.17 MB. These were measured before 0.5.0, whose reader limit checks make reading up to about 10% slower and allocate about 10% less.
- At 1,000,000 rows, SheetStream.NET with `Fastest` takes 4.8 s and allocates 275 MB; MiniExcel takes 5.3 s and 3.3 GB. The default setting takes 12.8 s for a file 17% smaller.
- The read benchmarks use a file SheetStream.NET wrote, and MiniExcel cannot bold the header, so its file carries slightly less formatting.

EPPlus and NPOI are pinned to their last plain open-source releases. Methodology, the wide-sheet scenarios and how to rerun everything: [benchmarks/README.md](benchmarks/README.md). Rerun it on your own hardware before relying on the numbers.

## When to choose something else

- You need to edit large workbooks in place. `ExcelPackage.Edit` holds the edited sheet in memory and is meant for small files. Use ClosedXML (MIT) or EPPlus.
- You need formulas evaluated in .NET, for example to read a calculated total without Excel. Use ClosedXML or EPPlus, which have calculation engines.
- You need sparklines, macros, encryption, or pivot tables with calculated fields, grouping or slicers. Not supported. See the [gap analysis](docs/gap-analysis-epplus-8.md) for the full list against EPPlus 8.
- You only map simple objects to and from sheets and want the smallest API. MiniExcel does that well.

## Migrating from EPPlus

[docs/migrating-from-epplus.md](docs/migrating-from-epplus.md) converts a real exporter line by line and has a quick-reference table.

## Building and testing

```bash
dotnet build
dotnet test
dotnet test tests/SheetStream.NetFx.Tests                          # .NET Framework 4.8 run of the netstandard2.0 build (Windows only)
dotnet run --project samples/SheetStream.Sample -- showcase        # Showcase.xlsx with every feature
dotnet run -c Release --project benchmarks/SheetStream.Benchmarks # comparison benchmarks
dotnet tool restore && dotnet docfx docs/docfx.json --serve        # documentation site on localhost
```

`samples/SheetStream.Sample.Web` is a minimal API that streams a 50,000-row export; see its README.

## Releasing

Bump `<Version>` in `src/SheetStream/SheetStream.csproj` (and in `src/SheetStream.AspNetCore/SheetStream.AspNetCore.csproj`), add a `CHANGELOG.md` entry, commit, then tag and push, for example `git tag v1.2.3` and `git push --tags`. The `Release` workflow stops if the tag and the csproj version disagree, and publishes through nuget.org Trusted Publishing, so no API key is stored in the repository.

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md). Questions go to [Discussions](https://github.com/P4Software/SheetStream.NET/discussions); security reports go through [SECURITY.md](SECURITY.md).

## Licence

MIT. See [LICENSE](LICENSE).
