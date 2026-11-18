# SheetStream.NET

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![NuGet](https://img.shields.io/nuget/v/SheetStream.NET.svg)](https://www.nuget.org/packages/SheetStream.NET) [![Downloads](https://img.shields.io/nuget/dt/SheetStream.NET.svg)](https://www.nuget.org/packages/SheetStream.NET) [![Build](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml/badge.svg)](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml) [![Coverage](https://img.shields.io/endpoint?url=https%3A%2F%2Fp4software.github.io%2FSheetStream.NET%2Fcoverage.json)](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml) [![OpenSSF Scorecard](https://img.shields.io/ossf-scorecard/github.com/P4Software/SheetStream.NET?label=OpenSSF%20Scorecard)](https://scorecard.dev/viewer/?uri=github.com/P4Software/SheetStream.NET) [![Docs](https://img.shields.io/badge/docs-p4software.github.io-0A5FAD)](https://p4software.github.io/SheetStream.NET/)

SheetStream.NET is a zero-dependency, MIT-licensed .NET library that creates, fills, edits and reads Excel (`.xlsx`) files. It is not related to the Node.js package of the same name.

P4 Software built it when EPPlus moved to a commercial licence. Rows are written straight into the zip archive as they are produced, so memory stays flat however many rows you export. It targets `netstandard2.0` and `net8.0`, which covers .NET Framework 4.7.2 and later, and .NET 8 and later. The namespace and assembly are `SheetStream`.

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
            .WriteCell($"Enterprise SKU-{i}")          // Text: shared or inline, chosen per column
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

Nothing is written until `Save()`, and an existing file is replaced only when it succeeds:

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

Dates come back as `DateTime`, elapsed times as `TimeSpan`, errors as `ExcelErrorValue`, and formula cells as their cached result. The full guide is [docs/reading-workbooks.md](docs/reading-workbooks.md). `ReadMergedRanges` and `ReadHyperlinks` read only the end of the sheet part, so they are fast on large sheets.

### Files from people you do not know

An `.xlsx` is a zip archive, so a small upload can unpack into gigabytes. `ExcelReader` rejects such files with an `InvalidDataException` that names the limit. The defaults sit far above anything Excel writes; change them with `ReaderOptions`:

```csharp
using var reader = ExcelReader.Open(upload, leaveOpen: false, new ReaderOptions { MaxCellTextLength = 32_767 }); // Excel's own cell limit
```

A stream that cannot seek is copied first: up to 64 MB in memory, beyond that to a temporary file that is deleted with the reader. `OpenTemplate` and `Edit` apply the same limits to the workbook they load, through `PackageOptions.SourceLimits`. See [SECURITY.md](SECURITY.md).

## Streaming rules

SheetStream.NET writes while you go, which is why memory stays flat. The cost is a few forward-only rules, each enforced with a clear exception (and, where the analyzer can prove it, a build warning):

1. One sheet at a time. Adding a worksheet seals the previous one.
2. Rows ascend. `Row(n)` must be the current row or a later one; gaps are fine.
3. Layout first. Column widths, frozen panes, tab colour, gridlines and fit-to-page are stored before the data, so set them before the first row. `AutoFitColumns()` and `Cells[...]` defer the header and lift this rule for that sheet.
4. A sheet uses `Row()` or `Cells[r, c]`, not both. `Cells[r, c]` gives EPPlus-style random access and buffers that one sheet in memory, so use it for small, irregular layouts.
5. Call `Save()`. Disposing a package does not save it: `Dispose` also runs when your code throws half-way, and a file that looks complete but is not would be worse than no file. A package that writes to a path (`Create`, `OpenTemplate`, `Edit`) builds the workbook in a temporary file beside the target and puts it in place only when `Save()` succeeds, so a failed export leaves an existing file as it was. The analyzer reports a package that is never saved (SS0004).

## Benchmarks

Writing 100,000 rows of 10 mixed, formatted columns. BenchmarkDotNet 0.15.8 on .NET 10.0.15, Windows Server 2025, an Azure `Standard_D8s_v6` virtual machine (Intel Xeon Platinum 8573C, 8 vCPUs, 32 GB) used for nothing else, measured on 2026-11-18 with three fresh processes per row (30 measured writes). Peak memory is the peak working set above an empty baseline process, from separate fresh-process runs. SheetStream.NET and SpreadCheetah each appear at both zip levels, so every row has a row from the other library at the same level to be read against.

| Library | Time | Allocated | Peak above baseline | File size |
|---|--:|--:|--:|--:|
| SheetStream.NET (`CompressionLevel.Fastest`) | 80.0 ms | 164 KB | 16.7 MB | 5,304 KB |
| SpreadCheetah 1.29.0 (its default, Fastest) | 106.8 ms | 12 KB | 15.0 MB | 5,296 KB |
| SheetStream.NET (default compression, Optimal) | 248.3 ms | 164 KB | 14.4 MB | 2,917 KB |
| SpreadCheetah 1.29.0 (Optimal) | 274.0 ms | 12 KB | 12.8 MB | 2,916 KB |
| MiniExcel 1.46.0 | 335.5 ms | 309 MB | 114.6 MB | 10,420 KB |
| NPOI 2.7.6 SXSSF | 1,582.3 ms | 636 MB | 130.9 MB | 5,518 KB |
| EPPlus 4.5.3.3 | 2,172.3 ms | 390 MB | 287.0 MB | 5,601 KB |
| ClosedXML 0.105.1 | 2,441.8 ms | 1,013 MB | 351.5 MB | 5,955 KB |
| NPOI 2.7.6 XSSF | 4,354.0 ms | 1,149 MB | 619.9 MB | 5,304 KB |

- **Against SpreadCheetah, at the same zip level:** SheetStream.NET is faster on this sheet (25% less time at `Fastest`, 9% at `Optimal`), and on wide sheets (1,000 × 1,000: 42% and 14% less), very wide sheets (1,000 × 10,000: 43% and 15%), a million rows (23% and 13%) and styled dates (14% and 3%). The two are level on a mix of repeated and unique text. SpreadCheetah is about 2% faster on its own benchmark shape, 100,000 rows of ten text columns in which every value is different, at the `Optimal` level. The same ratios came out of three complete runs.
- **Against MiniExcel:** SheetStream.NET at its default settings is faster in every scenario we measure except one: with every cell styled as a date, money or percentage (StyledDates), MiniExcel takes 301 ms against 341 ms at the default level (105 ms with `Fastest`). MiniExcel's files are 3.6 times larger on this sheet.
- **Allocation:** SpreadCheetah allocates the least (12 KB per write here), SheetStream.NET next (164 KB, the same for a million rows), then every other library by a wide margin.
- Reading the same 100,000-row file: SheetStream.NET 351 ms and 124 MB allocated, ExcelDataReader 693 ms and 249 MB, EPPlus 4.5.3.3 1,163 ms and 805 MB, MiniExcel 1,459 ms and 2,157 MB, ClosedXML 2,663 ms and 1,509 MB.
- At 1,000,000 rows, SheetStream.NET with `Fastest` takes 0.81 s, SpreadCheetah 1.05 s and MiniExcel 3.5 s; at the default level SheetStream.NET takes 2.5 s for a file 45% smaller.
- The read benchmarks use a file SheetStream.NET wrote, and MiniExcel cannot bold the header, so its file carries slightly less formatting.

These are the numbers of one machine; the ratios were the same in all three runs, the absolute times will differ on yours. EPPlus and NPOI are pinned to their last plain open-source releases. Every table, the raw output, the methodology and how to rerun everything, including the rows where SheetStream.NET is not first: [benchmarks/README.md](benchmarks/README.md). Rerun it on your own hardware before relying on the numbers.
## When to choose something else

- You need to edit large workbooks in place. `ExcelPackage.Edit` holds the edited sheet in memory and is meant for small files. Use ClosedXML (MIT) or EPPlus.
- You need formulas evaluated in .NET, for example to read a calculated total without Excel. Use ClosedXML or EPPlus, which have calculation engines.
- You need sparklines, macros, encryption, or pivot tables with calculated fields, grouping or slicers. Not supported. See the [gap analysis](docs/gap-analysis-epplus-8.md) for the full list against EPPlus 8.
- You only map simple objects to and from sheets and want the smallest API. MiniExcel does that well.
- Other tools read your files and need every cell's address. By default SheetStream.NET omits the 
 attribute on a cell that directly follows the previous one, which the Open XML specification allows and Excel reads; MiniExcel 1.46.0, NPOI 2.7.6's SAX reader and the Excel OleDb provider (ACE) misread such files. Set PackageOptions.WriteCellReferences = true, at the price of a larger file and a slower write.

## Migrating from EPPlus or ClosedXML

[docs/migrating-from-epplus.md](docs/migrating-from-epplus.md) converts a real exporter line by line and has a quick-reference table. [docs/migrating-from-closedxml.md](docs/migrating-from-closedxml.md) does the same for ClosedXML and says plainly what ClosedXML remains better at.

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
