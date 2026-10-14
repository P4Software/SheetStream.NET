# Changelog

All notable changes to SheetStream. Versions follow [Semantic Versioning](https://semver.org/).

## 0.4.0 (2026-10-14)

### Added

- **Edit small existing workbooks**: `ExcelPackage.Edit(path)` (in place, atomic replace) or `Edit(input, output)`, then `Worksheets.Edit(name)` loads the sheet into `Cells[...]`: read values, change cells, add rows, `Cells.Existing`, `Cells.Dimension`, `ExcelCell.Clear()`. Untouched cells, rows, formulas (shared and array formulas included), charts, tables and other sheets are written back unchanged. Guarded by `PackageOptions.MaxEditableCells` (200,000 by default); large files keep using streaming or templates.
- **Trimming and Native AOT**: the `net8.0` build is `IsTrimmable` and `IsAotCompatible`. A bundled **source generator** turns `[SheetRow] partial` types into compile-time mappers (`ISheetRowMapper<T>`, `SheetMappers`), so `LoadFromCollection<T>`, `LoadFromAsyncEnumerable<T>` and `ExcelReader.Read<T>` work without reflection or `Expression.Compile`; new overloads take a mapper explicitly. Reflection stays the fallback for unmarked types and is now annotated for the trimmer.
- **Roslyn analyzer** shipped inside the package: **SS0001** layout setting after rows were written, **SS0002** `Row()` and `Cells[]` mixed on one sheet, **SS0003** rows written out of order. Warnings at design time instead of exceptions at run time.
- **Reader**: `ReadMergedRanges` and `ReadHyperlinks` (issue #6) read a sheet's merged cells and links, external URLs resolved from the sheet relationships, without parsing the rows.
- **Benchmarks**: 1,000,000-row and 1,000 × 10,000-column write scenarios, and read benchmarks against MiniExcel, ExcelDataReader, ClosedXML, EPPlus and NPOI.
- **Web sample** (`samples/SheetStream.Sample.Web`): a minimal API with a download button, a 50,000-row streamed export and a `Cells[...]` report with charts; double-click `run.cmd`. A [DotNetFiddle](https://dotnetfiddle.net/BR59ju) runs the quick start in the browser.
- **Trust signals**: OpenSSF Scorecard workflow and badge, an SPDX SBOM and a signed build-provenance attestation attached to every release (`gh attestation verify <nupkg> --owner P4Software`), a coverage badge served from the docs site, and every GitHub Action pinned to a commit SHA.

### Changed

- The release workflow signs provenance in a job that only sees the finished packages; the build job stays read-only.

## 0.3.0 (2026-10-13)

### Added

- **Template filling**: `ExcelPackage.OpenTemplate(template, output)` and `Worksheets.Open(name)` stream rows into a workbook designed in Excel. Everything not touched is copied unchanged; tables, filters, data validations and conditional formats on the first data row stretch over the new rows; template styles are reusable with `TemplateStyle("D5")`; new styles append to the template's style table; the calculation chain is dropped and the workbook recalculates on open.
- **Async loading** (.NET 8+): `Worksheet.LoadFromAsyncEnumerable(...)` streams an `IAsyncEnumerable<T>` such as an EF Core query, with cancellation.
- **SheetStream.AspNetCore** package: `ExcelResults.Excel(...)` and `Results.Extensions.Excel(...)` return a streamed download from minimal APIs and controllers, with flat memory, accurate `Content-Length`, correct non-ASCII file names and no synchronous I/O.
- **Documentation site** at https://p4software.github.io/SheetStream/ with guides and the API reference generated from the XML comments.
- **Migration guide** from EPPlus, built on a real production exporter and compiled as a test.
- **Benchmarks** project comparing SheetStream with EPPlus 4.5.3.3, ClosedXML, MiniExcel and NPOI.
- Reader **fuzz tests**, code coverage in CI, `SECURITY.md` and Dependabot.

## 0.2.0 (2026-10-13)

Closes most of the gap with EPPlus while keeping the streaming, zero-dependency design.

### Added

- **Excel tables** with built-in styles, banding and filter buttons (`Worksheet.AddTable`).
- **Charts**: column, bar, line, area, pie, doughnut and scatter, with titles, axis titles, legends, data labels and series colours (`Worksheet.AddChart`).
- **Pictures**: PNG, JPEG and GIF, sized by pixels with the aspect ratio kept (`Worksheet.AddPicture`).
- **Conditional formatting**: cell value, expression, contains text, duplicates, unique, top / bottom, above / below average, colour scales, data bars and icon sets, with differential styles (`StyleSheet.CreateDifferential`).
- **Data validation**: lists, whole numbers, decimals, dates, text length and custom formulas, with prompts and error alerts.
- **Hyperlinks** to the web, e-mail and places in the workbook, using Excel's built-in Hyperlink style.
- **Comments** (notes), **rich text** cells and **formulas with cached values**.
- **Protection**: sheet and workbook protection with salted SHA-512 password hashes, plus locked / hidden-formula cell styles.
- **Print setup**: orientation, paper size, scale, fit to page, margins, header and footer, print titles, print area, gridlines and centering.
- **Layout**: AutoFit columns, tab colour, gridline visibility, zoom, right-to-left, row and column outlining, hidden rows, sheet visibility (hidden and very hidden).
- **Defined names** and **document properties** (title, author, company...).
- **Load helpers**: `LoadFromCollection` (with `[SheetColumn]` and `[DisplayName]`), `LoadFromDataTable` and `LoadFromDataReader`, optionally as a table.
- **Random-access cells**: `Worksheet.Cells[row, column]` for small, irregular layouts.
- **Streaming reader** (`SheetStream.Reading.ExcelReader`): `ReadRows`, `Read<T>` and `ReadDataTable`, with typed dates, durations and errors.
- More cell style options: indent, text rotation and shrink to fit.
- The NuGet package now carries the official P4 Software logo as its icon.
- Marker colours on line and scatter series; the reader rejects cell references beyond column XFD.

### Changed

- A workbook saved with no sheets now gets an empty "Sheet1", which Excel requires.
- `Dispose` no longer retries a save that already failed, so the original exception is not masked.
- Single-cell absolute references are written as `$B$2` rather than `$B$2:$B$2`.

## 0.1.0 (2026-10-13)

First release: streaming .xlsx writer with shared strings, styles (fonts, fills, borders, number formats, alignment), column widths, frozen panes, AutoFilter, merged cells, formulas and inline strings.
