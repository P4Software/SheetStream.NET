---
title: "Streaming Excel (.xlsx) library for .NET"
description: "A zero-dependency, MIT-licensed .NET library for creating and reading Excel (.xlsx) files. Rows stream straight into the file with flat memory."
---

# SheetStream.NET

A zero-dependency, MIT-licensed .NET library for creating and reading Excel (`.xlsx`) files. Rows stream straight into the file, so 10 million cells write in seconds with flat memory.

```bash
dotnet add package SheetStream.NET
dotnet add package SheetStream.AspNetCore   # optional: one-line Excel downloads
```

## Five-minute tour

```csharp
using SheetStream;
using SheetStream.Styles;

using var package = ExcelPackage.Create("orders.xlsx");
var workbook = package.Workbook;
var header = workbook.Styles.Create(o => { o.Font.Bold = true; o.Fill.BackgroundColor = "#0A5FAD"; o.Font.Color = "#FFFFFF"; });

var sheet = workbook.Worksheets.Add("Orders");
sheet.AutoFitColumns();
sheet.FreezePanes(rows: 1);
sheet.LoadFromCollection(orders, new LoadOptions { HeaderStyle = header, TableName = "Orders" });

var chart = sheet.AddChart(ChartType.Column, "H2:P20");
chart.AddSeries(values: "C2:C13", categories: "A2:A13", name: "=C1");

package.Save();
```

`Save()` is what writes the file. Disposing a package without it writes nothing: `Dispose` also runs when your code throws half-way through, and a file that opens but holds half the rows is worse than no file. `ExcelPackage.Create(path)` builds the workbook in a temporary file beside the target and puts it in place only when `Save()` succeeds, so a failed export never destroys the report that was there before. If you forget the call, the analyzer that ships in the package says so at build time ([SS0004](analyzers.md#ss0004)).

## Formulas and cached results

`row.WriteFormula("SUM(C2:C101)")` writes the formula and asks Excel to calculate on open. Viewers that do not calculate (file previews, Quick Look, many mobile and web viewers, `ExcelReader`) show such a cell blank unless the file also carries the result. Turn on `sheet.CacheFormulaResults()` before the first row and SheetStream computes the result itself when the formula is simple enough: numbers, references to cells already written on the same sheet, ranges, `+ - * /`, `%`, parentheses and `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `ROUND`, `ABS`. Anything else (other functions, other sheets, rows not yet written, a division by zero, a formula nested more than 64 levels deep) is written exactly as before, without a value; a wrong value is never written. The sheet keeps about 8 bytes per numeric cell to do this, which is why it is opt-in. When you already know the result, `WriteFormula(formula, cachedValue)` stores it directly.

## Large exports

Nothing needs switching on: rows are written as you produce them, so memory does not grow with the row count. Two things are held until the file is saved, the style table and the shared string table (see [How text is stored](#how-text-is-stored): with the default settings only the columns whose values repeat feed it). The shared string table stops growing at `PackageOptions.MaxSharedStringChars` (16,000,000 by default, a few tens of megabytes); after that, strings it has not seen are written inline in the sheet and strings already in it stay shared. An export of millions of unique strings therefore stays flat. `PackageOptions.CompressionLevel = CompressionLevel.Fastest` roughly halves the write time in exchange for larger files. How much larger depends on the runtime: about a quarter on .NET 8 and about 80% on .NET 10 in our measurements, because newer runtimes compress less at that level. To keep files small, a cell in the column right after the previous one is written without its address (`r="B2"`), which the file format allows. A few tools cannot read such cells: NPOI's event (SAX) reader, the Excel ADO / OleDb provider, Microsoft Spreadsheet Compare, and Open XML SDK code that reads `Cell.CellReference`. If you need one of them, set `PackageOptions.WriteCellReferences = true` and every cell is written with its address, at the cost of a larger file and a slower write.

## How text is stored

A text cell can be stored in the .xlsx in two ways: once in a workbook-wide list (the shared string table), with the cell holding only its number, which is what Excel itself writes; or inside the cell ("inline"). Excel, `ExcelReader` and every other reader show the same text either way. Only the file layout, the write speed and the memory differ. `PackageOptions.StringStorage` chooses:

| Value | What it does | Choose it when |
|---|---|---|
| `Auto` (default) | Decides column by column while the sheet is written. The first cells of each column are looked up in the table and counted: a column in which none of the first 128 cells repeats is judged unique at once, and any other is judged after 512 cells. A column whose values repeat (at least 5% of the sample) and are worth sharing (the repeated texts average 16 characters or more, or the sample holds 200 or more different values) keeps using the table. Every other column, such as ids, e-mail addresses, free text, or short values from a small set like a status or a country, is written inline from then on, without being hashed or remembered. A shared column that stops repeating (under 2% over 4,096 cells) is written inline from then on. A sheet of fewer than 128 rows is stored exactly as `Shared` stores it. | Nearly always. It is the fastest and the lightest on memory, and the file is as small as the other choices to within about 1% on the sheets we measured (100,000 rows of mixed text columns) |
| `Shared` | Every text goes through the table, which is what SheetStream did before this setting existed. | You want the file to be laid out the way Excel lays it out, or a sheet dominated by short repeated texts (hundreds of thousands of status, country or category cells) must open as quickly as possible in Excel: it opened about a fifth faster than the same sheet stored inline in our test, at the price of a slower write and a table held in memory until the save |
| `Inline` | Every text is written in its cell and no table is built. The same as `UseSharedStrings = false`, which still works and wins over this setting. | Memory must be flat and the data is known not to repeat |

The text of a column is judged once, from its first cells, so a column that starts as unique text and starts repeating further down stays inline; the file is then larger than `Shared` would make it, never larger than a writer that always writes inline. Sampling puts at most 512 texts per column into the table (128 for a column of unique text), and the table never grows past `PackageOptions.MaxSharedStringChars`. The bookkeeping is one small record per column that holds text (at most 16,384 of them, about 260 KB for the widest sheet Excel allows), whatever the number of rows.

## Where to go next

- [Migrating from EPPlus](migrating-from-epplus.md): a real exporter converted line by line.
- [Migrating from ClosedXML](migrating-from-closedxml.md): the same export in both libraries, and what each is better at.
- [Reading workbooks](reading-workbooks.md): rows, objects and DataTables, and the limits that make uploaded files safe to read.
- [Filling templates](templates.md): design in Excel, fill from code.
- [Editing existing workbooks](editing-workbooks.md): open a small workbook, change cells, save.
- [Pivot tables](pivot-tables.md): pivots over a table or range, refreshed when Excel opens the file.
- [Trimming, Native AOT and the source generator](aot-and-source-generator.md): object mapping without reflection.
- [Analyzers](analyzers.md): build-time warnings for the streaming rules.
- [Gap analysis versus EPPlus 8](gap-analysis-epplus-8.md): area by area, where SheetStream.NET is ahead, where it is not, and the order the gaps close.
- [A 98 KB spreadsheet that takes 400 MB](small-file-large-memory.md): how a small upload can exhaust a server's memory, what the 0.5.0 fix missed, and what changed in 0.6.0.
- [API reference](api/SheetStream.yml): every public type, generated from the source.
- [Benchmarks](https://github.com/P4Software/SheetStream.NET/tree/main/benchmarks): reproducible comparisons with EPPlus, ClosedXML, MiniExcel and NPOI.

## Design in one paragraph

SheetStream writes each worksheet directly into the zip archive as the rows arrive, so it never builds the workbook in memory. That buys flat memory and speed, and costs random access: rows are written in ascending order, one sheet at a time, and layout that Excel stores before the data (column widths, frozen panes) is declared first. Sheets that need random access can use `sheet.Cells[row, column]`, which buffers that one sheet. A workbook that exercises every feature is validated against the ECMA-376 schema in CI, and pivot tables are also checked by opening them in Excel itself (`tests/excel-check-pivot.ps1`).
