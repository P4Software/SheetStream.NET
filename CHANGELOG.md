# Changelog

All notable changes to SheetStream. Versions follow [Semantic Versioning](https://semver.org/).

## Unreleased

### Changed

- **Smaller sheet files: cells no longer repeat their address.** A cell that sits in the column right after the previous cell of its row is written without its `r="B2"` attribute, which the Open XML specification makes optional (a cell without one belongs in the next column; the first cell of a row belongs in column A). A cell that follows a gap, whether from `Skip()` or an empty cell with no style, still carries its address, so every value stays exactly where you put it. Excel, ExcelReader, `Worksheets.Edit` and `OpenTemplate` all read these files as before, and so do EPPlus, ClosedXML, MiniExcel, ExcelDataReader and NPOI's normal XSSF reader (all checked). A few tools do need the address on every cell: NPOI's XSSF event (SAX) reader, the Excel ADO / OleDb provider, Microsoft Spreadsheet Compare, and Open XML SDK code that reads `Cell.CellReference` (it is null for such a cell). If you or your users rely on one of them, set `PackageOptions.WriteCellReferences = true`: every cell is then written with its address, as before this change, at the price of a larger file and a slower write. The option is new and off by default.
- **Writing numbers creates no garbage.** Every number, style index and shared-string index written to a cell used to be turned into a new string first. They are now formatted straight into a reused buffer, so writing a million cells of numbers, dates and repeated text no longer allocates per cell (about 25 megabytes less garbage for a 100,000 by 10 sheet), which also means fewer garbage collections and a lower peak under load. The text in the file is unchanged, byte for byte. On .NET Framework (the `netstandard2.0` build) doubles, decimals, dates and durations still go through a string, because that runtime cannot format them into a buffer; integers, style numbers and text are free of it there too.

- **Faster cell writing.** Numbers, dates, durations, booleans, text (shared and inline) and styled empty cells are now put together as finished XML text in a reused buffer and handed to the XML writer in one piece per batch, instead of one element and one attribute at a time. The file is identical, byte for byte (checked against the previous build on a large set of cells and every character, including the escaping of `&`, `<` and `>`). Formulas, rich text and error cells are written as before. The speed-up depends on the machine and the data; the figures for this release come from the benchmark runs in `benchmarks/`, not from this note.

## 0.6.0 (2026-11-14)

A security and safety release. It closes, for every way of opening a workbook, the class of problem that 0.5.0 fixed only in part, and it changes when a workbook is written. Two changes can break existing code; they are listed first.

### Breaking changes

- **`Dispose` no longer saves.** A package writes its workbook only when you call `Save()`. `Dispose` also runs when your code throws half-way through, and until now that produced a file that opened normally and held half the rows. Code that relied on the `using` block to save now writes nothing; add `package.Save()` where the workbook is complete. The analyzer reports such code at build time (SS0004). `ExcelPackage.CreateBytes`, `CreateBytesAsync` and `ExcelResults.Excel` always saved explicitly and are unaffected.
- **A sheet too large to edit is refused with `InvalidDataException`**, not `InvalidOperationException`, so that one `catch` covers every limit on a workbook the library did not write. This is the `PackageOptions.MaxEditableCells` check in `Worksheets.Edit`.

### Changed

- **A package that writes to a path is all-or-nothing.** `ExcelPackage.Create(path)` and `OpenTemplate(template, outputPath)` now build the workbook in a temporary file beside the target and put it in place when `Save()` succeeds, as `Edit(path)` already did. A failed or abandoned export leaves whatever was at the path exactly as it was, where it used to leave a truncated file. The target does not exist until `Save()` returns.
- **`Save()` after a failure is refused** with an `InvalidOperationException` that says why, where it used to write a broken workbook: after a failed save, or after the source workbook of an edit or template could not be read.
- **`OpenTemplate` and `Edit` apply the reader's limits** to the workbook they load, through the new `PackageOptions.SourceLimits`. They were documented as being for files your application owns; they are now safe on a workbook a user uploaded.
- **`ReaderOptions.MaxLoadedPartBytes` now bounds memory, not only file size.** Besides capping the unpacked size of each part held in memory, the library adds up what those parts occupy once parsed (a string costs two bytes a character) and refuses a workbook when the total passes the limit. Two things that 0.5 accepted can now be refused at the default of 256 MB, each with a message naming the property to raise: a workbook with about three million or more distinct shared strings of 30 characters, and a `ReadDataTable` of a sheet with roughly 400,000 rows of ten numeric columns or more (a table that size takes a few hundred megabytes). `ReadRows` and `Read<T>` are not affected.
- **Parts are decoded as text before they reach the XML parser.** UTF-8 and UTF-16 are detected from the byte order mark or the first bytes, and an encoding named in the XML declaration is honoured when the runtime provides it. Bytes that are not valid in the part's encoding are refused as damage.
- Comments and processing instructions inside the XML of a template sheet that is filled are no longer copied to the output. Excel writes neither.

### Security

Versions up to and including 0.5.1 can be made to exhaust memory, or to run for minutes, by a small crafted `.xlsx`. 0.5.0 fixed the case that was reported (one very long cell) and missed others of the same class. All of the following are fixed in this release, for `ExcelReader`, `OpenTemplate` and `Edit` alike. If your application opens workbooks from people you do not know, upgrade.

- **A CDATA section** was buffered whole by the XML parser before any limit could see it: a 98 KB file reached 418 MB in the reader.
- **A part encoded as UTF-16** could hide a long attribute from the length check, which counted bytes: characters such as U+3C3C look like markup when seen as bytes.
- **A row of cells without references** had no width limit: a 524 KB file made the reader hold 313 MB for one row.
- **Many different element or attribute names** grew the parser's name table without limit, in a streamed sheet as well.
- **A tag with a very large number of attributes** cost the parser hundreds of megabytes while staying under the length limit that was meant for text.
- **`OpenTemplate` and `Edit` had no limits at all**: a 429 KB file reached 5.6 GB in `Edit`, and template filling of a 16 KB file of deeply nested elements was still running after five minutes when it was stopped.
- **A template row with a number far outside Excel's range** made template filling allocate in proportion to the number.
- **A workbook whose relationships name a part that is not in the file** threw `NullReferenceException` in template and edit mode. It is now reported as damage.
- **A streamed sheet could be any size**, so a small upload could cost minutes of processing. The new `ReaderOptions.MaxStreamedPartBytes` (default 4 GB) bounds it.
- **A zip directory can point many names at one compressed block.** Saving an edit or a template copied every one: a 259 KB file made `Save()` unpack and repack 10 GB, and the format allows far more. The parts copied unchanged are now capped in total by `MaxStreamedPartBytes`.
- **`ReadDataTable` had no bound.** A table is rows times columns, so 3,000 rows with one cell each in the last column (15 KB on disk) built a table of 1 GB. What the table holds is now counted against `MaxLoadedPartBytes`.
- **Three things took time out of all proportion to the file.** A relationship target made of `./` repeated was resolved in quadratic time (a 6 KB file took 43 seconds to open); a very long number format was inspected again for every cell format, and in edit mode for every cell (a 10 KB file took 18 seconds); and a `DataTable` header of identical names was made unique by counting up from 2 for each one.
- **A formula from a file could overflow the stack.** With `CacheFormulaResults()`, a formula nested a few thousand levels deep ended the process, which no `catch` can stop. The evaluator now leaves a formula uncached when it is nested deeper than 64 levels or longer than 8,192 characters, Excel's own limits.
- **`CacheFormulaResults()` on an edited sheet could take 1.6 GB** when the sheet's cells were scattered one to a page of the value store. The store now stops at 256 MB, after which formulas go uncached.
- **Part names from the source workbook were copied into the output unchecked**, including `../../name` and absolute paths, which can place a file outside the folder a careless tool unpacks into, and the same name twice. Such a workbook is now refused.

Measured with the default limits on seventeen kinds of file built to exhaust memory, each about 100 MB unpacked, through reading, editing, template filling and template copying: the process now peaks at 65 MB or less, where 0.5.1 reached up to 5.9 GB. The one exception is a workbook whose shared strings part is genuinely 100 MB, which uses 250 MB, inside the 256 MB the default limit allows.

### Fixed

- **`Worksheets.Edit` counted only every other cell against `MaxEditableCells`.** A sheet of between one and two times the limit was refused late, after part of it had been loaded, and saving the package afterwards wrote that sheet with cells missing. The count is now exact and made before anything is loaded.
- **Workbooks in the 1904 date system written by Excel were read four years and a day off.** Excel writes a second `workbookPr` element in an extension list, and that one switched the date system back to 1900. Dates in such workbooks are now right in `ExcelReader` and in `Worksheets.Edit`. Found by comparing the reader with Excel on files Excel wrote.
- **A `Save()` that could not put the file in place could not be retried.** When the target was open in Excel, `Save()` threw, and a second call returned without doing anything, so the workbook was lost while the message said to try again. The finished workbook now stays in its temporary file and the next `Save()` puts it in place. `IsSaved` is true only once the file is where it was asked to be.
- `ReadDataTable` threw `DuplicateNameException` when a data row was wider than the header and the header already used the letter of the added column as a text.
- `ExcelRow.GetDateTime` threw `ArgumentException` for a number outside the range of dates; it is now the `FormatException` every other failed conversion gives.
- In `Worksheets.Edit`, a date cell read as a plain number when its format id shared its code with an earlier format id.
- Strings added to a workbook whose shared strings part has an `si` element nested in an extension pointed one place too far.
- A target with a long file name failed in `Create`, because the temporary file was named after it plus 37 characters.
- A reader or package that fails to open releases the stream it was given when it was told to own it.
- **On .NET Framework, a damaged zip could surface as `ArgumentException` or `IOException`.** That runtime reads the zip directory on first use and reports some damage with those types. The directory is now read when the file is opened, on every runtime, and damage is reported as `InvalidDataException`.
- On .NET Framework, a number read as text had up to 17 digits where 16 are enough (`0.33333333333333331`). `GetString` and `Read<T>` into a string now give the shortest text that reads back, the same on every runtime.
- Filling a template whose range lists hold very many tokens that are not ranges took seconds per list, an exception per token. Ranges read from a file are now parsed without exceptions.
- `Cells.Dimension` walked every cell on each call, so appending rows with `Cells[Dimension.LastRow + 1, 1]` was quadratic. `TemplateStyle` re-read the sheet on every call; it now remembers its answers.
- A template with a number format id near the top of the integer range made the next new format get a negative id, an invalid file.
- A relationship that points at a missing sheet part, a workbook with two sheets of one name, a part that consists of an empty root element and a part whose root is the wrong element are all reported as a damaged file (`InvalidDataException`). They used to end in `ArgumentException` or `InvalidOperationException`.

### Added

- `PackageOptions.SourceLimits` and `ReaderOptions.MaxStreamedPartBytes` (see above).
- **Analyzer rule SS0004**, "Package is never saved": a package created in a method that never calls `Save()` on it.
- Passing a `CellStyle` as a cell value throws `ArgumentException`. `WriteCells(style, "a", "b")` compiles, binds to the overload that takes only values, and used to write the style's type name into the first cell; the message now says to write `WriteCells(style, new[] { "a", "b" })`.
- A guide for [migrating from ClosedXML](docs/migrating-from-closedxml.md).
- Fixed limits that need no setting because no real workbook comes near them: 16,384 distinct XML names per part, 16,384 cells per row, a tag no longer than a cell's text limit, row numbers within Excel's range in a sheet that is edited or filled.

### Testing

- About sixty kinds of hostile workbook are run through every way of opening one, and the outcome of each combination is recorded and compared on every build.
- The same combinations run in a separate process whose heap is capped at 64 MB by the runtime, which proves peak memory in a way an in-process test cannot. They pass at 40 MB.
- The template and edit code is fuzz-tested like the reader, from library-written workbooks and from workbooks Excel wrote. Deep runs before this release fed several hundred thousand randomly damaged files through every way of opening a workbook, on a dozen seeds.
- **A separate adversarial review pass was run on the work before release**, with its own exploit harness and fuzzer, and found thirteen problems the tests above had missed: unbounded time, the zip container, and what callers collect. Each is fixed and has a regression test (`ReviewFindingsTests`), and the review's own repros and fuzzer were rerun against the fixes.
- The limits are tested on .NET Framework 4.8 as well.
- **A corpus of workbooks written by Excel** (`tests/corpus/excel`), each with Excel's own account of every cell: its value and type, the date it displays, merges, links, sheet order and visibility. The reader is compared with it cell by cell, and again after a template copy and after an edit, so data lost in a round trip shows up as a difference from Excel.
- `tests/excel-check.ps1` opens workbooks in an installed Excel and fails on any that Excel has to repair. It is run before a release on everything the samples write and on every corpus workbook after a copy, an edit and a fill.

## 0.5.1 (2026-11-13)

No change to the library's code or behaviour; this release updates what the packages carry and show.

### Changed

- **README rewritten** for both packages. The package pages on nuget.org now show a shorter README with plain headings, exact benchmark figures and links to the docs, and `SheetStream.AspNetCore` has a complete page of its own (what it solves, how the response behaves, options).
- The OpenSSF Scorecard badge is served from a domain nuget.org allows, which removes the "images are not displayed as they are not from trusted domains" banner from the package page.

### Security

- The reader memory issue fixed in 0.5.0 is published as security advisory [GHSA-rvj5-c489-67jr](https://github.com/P4Software/SheetStream.NET/security/advisories/GHSA-rvj5-c489-67jr). It affects `SheetStream.NET` before 0.5.0 and the retired `SheetStream` package 0.2.0 to 0.3.0. If you read workbooks from untrusted users, upgrade.

### Build

- Package versions for the build, tests and samples are declared once in `Directory.Packages.props`. The web sample uses the published `SheetStream.NET` package.
## 0.5.0 (2026-11-13)

A hardening release: it fixes defects found in a review of 0.4.0. Two behaviours change; both are listed first.

### Changed

- **In-place edits commit only on `Save()`.** `ExcelPackage.Edit(path)` used to save when the package was disposed, like every other package. Dispose also runs when your code throws half-way through its changes, so a half-finished edit replaced the original file. Now leaving the `using` block without calling `Save()` discards the changes and leaves the file untouched. **If you relied on dispose to save an in-place edit, add `package.Save()`.** Creating a file, filling a template and `Edit(input, output)` still save on dispose.
- **The shared string table has a memory budget.** `PackageOptions.MaxSharedStringChars` (default 16,000,000) caps what the table holds until the save. Once it is full, strings it has not seen are written inline and strings already in it stay shared, so repeated values still de-duplicate while an export of millions of unique strings no longer grows memory without limit (it used to retain about 122 MB per million short unique strings). Files are identical for exports under the budget.

### Fixed

- **The reader could be made to exhaust memory by a small file.** A 192 KB workbook holding one 200-million-character cell drove a reading process to 1.5 GB. `ExcelReader` now reads cell text in chunks against limits and rejects such a file with an `InvalidDataException` as soon as a value passes its limit; the same file now peaks at 32 MB. New `ReaderOptions` (`MaxCellTextLength`, `MaxRowTextLength`, `MaxLoadedPartBytes`) on new `ExcelReader.Open` / `OpenAsync` overloads; oversized tags and absurdly deep XML are rejected too, and merged-range and hyperlink lists are capped. The defaults sit far above anything Excel writes.
- **`ExcelReader.Open(Stream)` buffered a non-seekable upload entirely in memory.** It now spills to a self-deleting temporary file above 64 MB, as `OpenAsync` already did.
- **A duration-formatted cell holding `NaN` threw `ArgumentException`**, outside the exception types the reader promises for damaged files. `NaN` and infinities now read as text, and durations or dates too large for their type as the raw number.
- **Cached `ROUND` results disagreed with Excel** (`ROUND(1.005, 2)` cached 1, Excel shows 1.01; `ROUND(0.285, 2)` cached 0.28, Excel shows 0.29). The evaluator now rounds the 15-significant-digit decimal value the way Excel does; verified against Excel 16 on 4,000 generated cases. A number too large to round exactly gets no cached value.
- **Dates before 1 March 1900 were written one day late.** Excel's serials differ from OLE dates there (it counts a 29 February 1900 that never existed): 1 January 1900 is now written as serial 1, and the reader applies the same rule in reverse.

### Known limits

- `ExcelPackage.OpenTemplate` and `ExcelPackage.Edit` still load the whole source workbook into memory and do not apply `ReaderOptions`; they are for workbooks your application owns. See `SECURITY.md`.
- Reading is up to about 10% slower than 0.4.0 on the 100,000-row benchmark file (and allocates about 10% less), the cost of the limit checks. The read benchmark tables predate this release.

### Added

- A `.NET Framework 4.8` test project (`tests/SheetStream.NetFx.Tests`) and a Windows CI job, so the netstandard2.0 build and Windows are covered by CI; the main tests now also run on .NET 8.
- Packing checks the public API against the previous release (`EnablePackageValidation`), so a breaking change cannot ship by accident.
- `global.json` pins the .NET 10 SDK.

## 0.4.0 (2026-10-14)

### Changed

- **Package renamed to `SheetStream.NET`** (the repository is now `P4Software/SheetStream.NET`, the docs live at https://p4software.github.io/SheetStream.NET/). The C# namespace stays `SheetStream` and `SheetStream.AspNetCore` keeps its name, so upgrading is one package reference: `dotnet remove package SheetStream` then `dotnet add package SheetStream.NET`. The old `SheetStream` package stays on nuget.org at 0.3.0, marked as renamed.

### Added

- **Edit small existing workbooks**: `ExcelPackage.Edit(path)` (in place, atomic replace) or `Edit(input, output)`, then `Worksheets.Edit(name)` loads the sheet into `Cells[...]`: read values, change cells, add rows, `Cells.Existing`, `Cells.Dimension`, `ExcelCell.Clear()`. Untouched cells, rows, formulas (shared and array formulas included), charts, tables and other sheets are written back unchanged. Guarded by `PackageOptions.MaxEditableCells` (200,000 by default); large files keep using streaming or templates.
- **Trimming and Native AOT**: the `net8.0` build is `IsTrimmable` and `IsAotCompatible`. A bundled **source generator** turns `[SheetRow] partial` types into compile-time mappers (`ISheetRowMapper<T>`, `SheetMappers`), so `LoadFromCollection<T>`, `LoadFromAsyncEnumerable<T>` and `ExcelReader.Read<T>` work without reflection or `Expression.Compile`; new overloads take a mapper explicitly. Reflection stays the fallback for unmarked types and is now annotated for the trimmer.
- **Roslyn analyzer** shipped inside the package: **SS0001** layout setting after rows were written, **SS0002** `Row()` and `Cells[]` mixed on one sheet, **SS0003** rows written out of order. Warnings at design time instead of exceptions at run time.
- **Pivot tables**: `Worksheet.AddPivotTable` over a table or range written earlier, with row, column, filter and data fields (Sum, Count, Average, Max, Min and more), styles and grand totals. The cache is marked refresh-on-open, so Excel computes the pivot when the file opens; verified against Excel 16 through COM automation (`tests/excel-check-pivot.ps1`).
- **Cached formula results**: `Worksheet.CacheFormulaResults()` opts a sheet into a small evaluator (`+ - * / %`, `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `ROUND`, `ABS` over cells already written) so viewers that do not calculate show totals instead of blanks. Anything it cannot prove is written without a value, as before.
- **Async reading** (.NET 8+): `ExcelReader.OpenAsync`, `ReadRowsAsync`, `ReadAsync<T>` and `ReadDataTableAsync`, with cancellation between rows; non-seekable uploads are buffered asynchronously and spill to a temp file above 64 MB.
- **From the first production migration (LogisticaHQ)**: `Worksheet.LoadFromArrays` replaces EPPlus' `Cells.LoadFromArrays` one for one; `ExcelPackage.CreateBytesAsync` lets a builder await a query per sheet; the migration guide gained "Things EPPlus did for you silently" (the code-page encoding provider that ExcelDataReader imports depend on, AutoFit timing, default date formats).
- **AutoFit** measures columns with a Calibri 11 glyph table calibrated against Excel's own AutoFit (bold, font size, formatted numbers and dates as displayed) and Excel's column-width formula, instead of counting characters; `Worksheet.AutoFitColumns(Func<string, bool, double> measure, ...)` and `PackageOptions.TextMeasurer` accept real font metrics in digit units.
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
- **Documentation site** at https://p4software.github.io/SheetStream.NET/ with guides and the API reference generated from the XML comments.
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
