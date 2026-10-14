# Gap analysis: SheetStream 0.4.0 versus EPPlus 8

*Written 2026-10-14 against EPPlus 8.7.1 (NuGet, 2026-09-22) and SheetStream main at `2bcb8fa` (0.4.0, unreleased). EPPlus facts come from EPPlus Software's own feature page, wiki, API reference, release notes and licence pages; the source list is at the end. SheetStream facts come from the repository and its tests. Nothing here is a benchmark of EPPlus 8 itself: its licence does not allow P4 Software to run it, so timings use the last LGPL release, 4.5.3.3.*

The goal set for SheetStream is **as good as or better than EPPlus 8 in every aspect**. This document says, area by area, where that is already true, where it is not, how big each gap is, and in what order to close them. Effort is given in agent-days at the pace of the 0.4.0 work (roughly one of these days per feature of the size of pivot tables or edit mode).

## 1. Summary

| Area | EPPlus 8 | SheetStream 0.4.0 | Verdict |
|---|---|---|---|
| Licence, dependencies, platforms | Commercial, US$569/dev/yr, 5 dependencies, 16 MB package | MIT, zero dependencies, 226 KB, Native AOT, bundled analyzer | **SheetStream better** |
| Performance and memory | Whole workbook in memory; official: 100k × 200 cells in 29 s | Streaming; 100k × 10 in 1.0 s, 27 MB allocated; 1M rows at 200 MB peak | **SheetStream better** |
| Writing cells, rows, layout | Random access anywhere | Streaming rows, plus `Cells[]` per sheet | Parity for exports; gap for in-place edits |
| Reading | Load whole package; typed values; `Text` | Streaming, async, fuzz-tested; typed values | **SheetStream better** for imports; gaps: `.xls`, CSV, formatted `Text`, reading definitions |
| Templates and editing | Open and edit any workbook | Template filling (streaming); edit small workbooks | Gap: large edits, insert/delete rows and columns, copy ranges and sheets |
| Styling | Named styles, gradients, themes, 45 conditional-format rules | Cell styles, differential styles, 15 conditional-format rules | Gap: medium |
| Data loading and export | 7 loaders, CSV/text, HTML, JSON, `ToDataTable`, `ToCollection` | 6 loaders incl. async enumerable and AOT mapper; readers | Gap: CSV, HTML, JSON export; range-to-object |
| Tables | Totals, calculated columns, filters, sorting, slicers, custom styles | Tables with built-in styles and filter buttons | Gap: medium |
| Pivot tables | Grouping, calculated fields, slicers, Show Values As, calculation, pivot charts | Fields, functions, styles, refresh on open | Gap: medium-large |
| Charts and drawings | 80+ chart types, styles, sparklines, shapes, controls, OLE, in-cell pictures | 15 chart types, pictures | Gap: large |
| Formulas | Calculation engine, 478 functions, dynamic arrays, LAMBDA | Written as text; cached results for simple aggregates; Excel calculates on open | Gap: largest |
| Validation, filtering, sorting | 7 validation types, filter criteria, range and table sorting | 6 validation types, AutoFilter | Gap: small-medium |
| Printing | Page setup, breaks, header/footer images, odd/even pages | Page setup, print titles and area, text headers/footers | Gap: small |
| Security | Encryption (read/write), digital signatures, VBA signing, sensitivity labels | Sheet/workbook protection hashes | Gap: medium (encryption); signatures and labels are niche |
| Tooling and support | Commercial support, 132 wiki pages, large sample repos | Docs site, analyzer, source generator, fiddle, GitHub Discussions | Gap: commercial support (by design); samples |

**Counting the work:** closing every gap that a business application could plausibly hit ("business parity", phases 1 to 5 below) is about **55 to 70 agent-days**. Full parity including a formula engine with EPPlus' 478 functions, every chart type and the drawing model is **130+ agent-days**, and some items should not be chased at all (section 4).

## 2. Where SheetStream is already better

These are the reasons to choose SheetStream today, and they must be protected while closing the gaps (every new feature keeps zero dependencies, the streaming path allocation-free, and the package small).

- **Licence and cost.** MIT versus US$569 per developer per year (subscription, 1 to 4 developers) or US$899 perpetual, with a licence key that must be set at startup and a SaaS clause that needs a custom licence. SheetStream has no key, no seat count and no SaaS clause.
- **Dependencies and size.** Zero package dependencies on every target; 226 KB assembly. EPPlus 8 brings `EPPlus.Interfaces`, `Microsoft.Extensions.Configuration.Json`, `Microsoft.IO.RecyclableMemoryStream`, `System.Security.Cryptography.Pkcs` and `System.Security.Cryptography.Xml`, in a 16 MB package.
- **Memory and speed for exports.** Writing 100,000 × 10 formatted cells: 1.0 s and 27 MB allocated (0.53 s with fastest compression); EPPlus 4.5.3.3 takes 2.3 s and 390 MB; EPPlus' own 8.6 figure for 100,000 × 200 cells is 29 s with the whole workbook in memory. One million rows write with a 200 MB peak working set including the output buffer. A 10,000-column sheet is the same cost as a 10,000-row one.
- **Memory for imports.** `ExcelReader` streams rows (2.7 MB above baseline for a 100,000-row file); EPPlus loads the package.
- **Native AOT and trimming.** `IsAotCompatible` and `IsTrimmable` on `net8.0`, with a source generator that removes reflection from object mapping. EPPlus has no AOT or trimming claim.
- **Design-time safety.** A Roslyn analyzer in the package turns the three streaming mistakes into IDE warnings. EPPlus validates nothing at design time, and by its own docs does not validate formulas or number formats on write either.
- **ASP.NET Core.** `ExcelResults.Excel` streams a download with accurate `Content-Length` and no synchronous I/O; `LoadFromAsyncEnumerable` feeds an EF Core query straight into a sheet.
- **Supply chain.** SHA-pinned actions, SBOM and signed build provenance on every release, Trusted Publishing, OpenSSF Scorecard, fuzz tests. EPPlus code-signs its assemblies with a certificate (which SheetStream does not, see 3.1).
- **Template filling at scale.** Streaming rows into an Excel-designed template keeps memory flat; EPPlus loads the template.

## 3. Gaps, by area

Each table lists what EPPlus 8 has, what SheetStream has, and the plan. "Business" marks features a line-of-business export or import is likely to hit; "Niche" marks features few applications use.

### 3.1 Packaging, licensing, platforms

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Target frameworks | net35, net462, netstandard2.0/2.1, net8/9/10 | netstandard2.0, net8.0 (covers .NET Framework 4.7.2+, .NET 8/9/10) | None in practice; net35 is not a goal | — |
| Code-signed assemblies | GlobalSign certificate | Signed build provenance (Sigstore), repository-signed by nuget.org | Small, Business (some procurement checklists ask for Authenticode) | Buy a code-signing certificate (~US$200/yr) and sign in the release job: 0.5 |
| Commercial support with SLA | 48-hour response, CET | GitHub issues and Discussions | By design | Offer a paid support option from P4 Software if demand appears |
| Configuration via appsettings | Licence only | Nothing to configure | None | — |

### 3.2 Workbook and file modes

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Create, save, streams, byte arrays, async save | Yes | Yes (`Create`, `CreateBytes`, `CreateBytesAsync`, ASP.NET result) | None | — |
| Open and edit any workbook in place | Yes (whole package in memory) | Small workbooks (`Edit`, 200,000 cells per sheet by default); large ones as templates | **Medium, Business** | Raise the practical limit by keeping untouched rows as raw XML and only materialising touched rows: 3. True large-file editing by streaming rewrite with a row-transform API: 4 |
| Insert and delete rows, columns and ranges with reference shifting | Yes, adjusts formulas, tables, pivots, validations, drawings | No | **Medium, Business** | In edit mode: shift cells, merges, tables, validations, conditional formats and formulas on the same sheet: 3 |
| Copy ranges and worksheets (with styles, across workbooks) | Yes, many options | No | Medium, Business | Copy range within a sheet and copy sheet within a workbook in edit mode: 2 |
| Password-protected files (encryption), read and write | Standard and Agile, AES 128/192/256 | No | **Medium, Business** (users upload encrypted files) | Read (Agile and Standard, AES/SHA with .NET inbox crypto): 2.5; write: 2 |
| Digital signatures, signature lines | Yes (8.0) | No | Niche | net8.0-only implementation with `System.Security.Cryptography.Xml` from the shared framework: 3. Defer |
| VBA: keep macros in `.xlsm` | Create, sign, protect | Untested (edit mode copies unknown parts) | Small, Business | Test and document `.xlsm` round-trip in edit and template mode; keep macro-enabled content types: 0.5. Creating modules: 3, defer |
| CSV / text import and export | `LoadFromText`, `SaveToText`, fixed width | No | **Small, Business** | `ExcelReader.OpenCsv` / `ReadCsv` and `Worksheet.LoadFromCsv`, `ExcelReader.ToCsv`: 1.5 |
| Save as template (.xltx) | Yes (8.7) | No | Small | Content-type switch at save: 0.3 |
| OLE objects, external connections and query tables, sensitivity labels | Yes (8.0 to 8.3) | No | Niche; labels need the Microsoft SDK | Non-goal (section 4) |
| Strict OOXML | Not supported | Not supported | None | — |

### 3.3 Worksheets and cells

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Random access, multi-area ranges, LINQ over cells | Yes | `Cells[r, c]` and `Cells["A1"]` per sheet; `Cells.Existing` | Small | Range object (`Cells["A1:D10"]`) with `Value`, `Style`, `Formula` setters over the buffer: 1 |
| Merged cells, hyperlinks, comments, rich text | Yes | Yes (reader returns merges and links) | None | — |
| Threaded comments | Yes | No (legacy notes only) | Small, Niche | 1.5 |
| AutoFit columns | Font-measured, pluggable `ITextMeasurer`, bundled fonts | Character count (width-table version in progress) | Small, Business | Width table for Calibri 11 plus bold factor and a measurer hook: in progress (0.4.0) |
| AutoFit rows | Yes | No | Small | Estimate wrapped-text height from the width table: 1 |
| Freeze panes | Yes | Yes | None | — |
| Split panes | Yes | No | Niche | 0.3 |
| Sheet views (zoom, gridlines, headers, RTL, page layout view, selected cell) | Yes | Zoom, gridlines, RTL | Small | Active cell / selection, page-layout view: 0.3 |
| Outline grouping with summary position, collapse APIs | Yes | Outline levels and collapsed flag | Small | `OutlineSummaryBelow/Right`: 0.2 |
| Row and column styles, default heights/widths | Yes | Row style, column width/style before data | Small | Default row height and column width: 0.2 |
| Sheet protection with 16 allow flags, protected ranges | Yes | Protection with allow options | Small | Protected ranges (`protectedRange` with password): 0.5 |
| Defined names (ranges, formulas, values), fill ranges, take/skip | Yes | Defined names | API style; none for files | — |
| Add, remove, reorder, hide sheets | Yes | Add, hide; order is write order | Small | Reorder at save (workbook part lists sheets, parts keep numbers): 0.5 |

### 3.4 Styling

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Fonts, fills, borders, number formats, alignment, protection, indent, rotation, shrink | Yes | Yes | None | — |
| Named styles (`cellStyles`) | Yes | No | Small, Business (corporate style sheets) | 1 |
| Gradient and pattern fills | Yes | Solid only | Small | 0.5 |
| Themes and theme colours | Load `.thmx`, modify colour scheme, theme colours in styles | No (fixed default theme) | Medium | Theme colours in styles and conditional formats (`theme`/`tint`): 1; load a theme from a template: 1 |
| Conditional formatting rules | 45 types incl. date periods, blanks/errors, begins/ends with, custom icon sets, data-bar options | 15 types: cell value, expression, contains text, duplicates, unique, top/bottom N, above/below average, 2- and 3-colour scales, data bars, icon sets | Medium, Business | Date periods, blanks/errors, begins/ends/not-contains, std-dev, percent variants, data-bar negative/axis options, custom icon sets: 2 |
| In-cell checkbox style | Yes (8) | No | Niche | 0.3 |
| Custom table, pivot and slicer styles | Yes | Built-in styles only | Small | Custom table style definitions: 1.5 |
| Drawing styling (colour manager, effects, 3D) | Yes | Series and marker colours | Part of charts (3.7) | — |

### 3.5 Data loading and export

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| `LoadFromCollection` with attributes | Yes; `EpplusTable` (totals, autofit), `EpplusTableColumn` (order, header, format, totals), formula columns, nested columns, header parsing (camel case to spaces) | Yes; `[SheetColumn]` (name, order, format, ignore), `[DisplayName]`, header style, table | Small, Business | Totals row options, formula columns, nested objects, camel-case headers in `LoadOptions`: 1.5 |
| `LoadFromDataTable`, `LoadFromDataReader` (+async), `LoadFromArrays`, `LoadFromDictionaries` | Yes | Yes (reader async via `LoadFromAsyncEnumerable`), dictionaries via `LoadFromCollection` | None | `LoadFromDataReaderAsync` with `DbDataReader`: 0.3 |
| Transpose on load | Yes | No | Niche | 0.5 |
| `LoadFromText` / CSV import | Yes, with formats and fixed width | No | Small, Business | See 3.2 CSV: shared |
| `ToDataTable`, `ToCollection<T>` from a range or table, with mappings and failure strategies | Yes | From a sheet only (`ReadDataTable`, `Read<T>`) | Small, Business | Range and table scope (`ReadRows(sheet, "A1:D100")`, header row / data start row options), lenient header matching (ignore case, spaces, underscores), conversion-failure strategy (throw, skip, null): 1 |
| CSV / text export (`ToText`, `SaveToText`) | Yes | No | Small, Business | 0.5 (with CSV import) |
| HTML export with CSS, conditional formats, accessibility | Yes | No | Medium, occasional | 3; defer behind CSV/JSON |
| JSON export | Yes | No | Small | 1 |
| `GetValue<T>` on cells | Yes | `ExcelRow.GetString/GetDouble/GetDateTime`; `Read<T>` | Small | `GetValue<T>` on `ExcelRow` and `ExcelCell`: 0.3 |

### 3.6 Tables and pivot tables

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Tables with built-in styles, filter buttons, banding | Yes | Yes | None | — |
| Totals row (functions, formulas), calculated columns, structured references | Yes | No | Small, Business | Totals row and calculated column formulas on `ExcelTable`: 1 |
| Table filter criteria written to the file (value, custom, top-10, dynamic, date group), `ApplyFilter` hiding rows | Yes | Filter buttons only | Medium | Write filter criteria for table and sheet AutoFilter; apply by hiding rows at write time: 1.5 |
| Table sorting, custom lists | Yes | No | Small | Sort rows before writing is the streaming answer; write `sortState`: 0.5 |
| Table slicers | Yes | No | Medium, occasional | With pivot slicers: 3 |
| `DataRows` API over a table | Yes (8.0) | Edit mode `Cells` | API style | — |
| Pivot: fields, 12 functions, layout, grand totals, styles | Yes | Yes (11 functions; `None` missing) | None | — |
| Pivot cache with data (viewers show values without Excel) | Yes (`SaveData`) | Refresh on open only | Medium, Business (previews, Google Sheets, LibreOffice) | Write cache records and shared items from the source rows (needs the source sheet's values: buffer the source table's columns while streaming, opt-in): 2 |
| Date and numeric grouping | Yes | No | Medium, Business | 2 |
| Calculated fields, Show Values As (14 options) | Yes | No | Medium | 1.5 |
| Pivot filters (item, label, value, date, top-10), sorting, page fields | Partly (page fields) | No | Medium | 2 |
| Pivot slicers and timelines | Slicers yes | No | Medium | 3 (with table slicers) |
| Pivot calculation in .NET (`Calculate`, `GetPivotData`) | Yes (7.2) | No | Large | Depends on the formula engine (3.8) |
| Pivot charts | Yes | No | Medium | After chart parity: 1.5 |
| Pivot-area styling, conditional formatting on pivots | Yes | No | Niche | Defer |

### 3.7 Charts and drawings

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Chart types | 80+ `eChartType` members: all 2D and 3D column/bar/line/area/pie, cone/cylinder/pyramid variants, doughnut, scatter (5), bubble, radar (3), stock (4), surface (4), box & whisker, funnel, histogram/Pareto, sunburst, treemap, waterfall, region map | 15: column (3), bar (3), line (3), area (2), pie, doughnut, scatter (2) | **Large, Business** for 2D; Niche for 3D and extended | Phase A, 2D parity: bubble, radar, stock, exploded pie/doughnut, pie-of-pie/bar-of-pie, scatter smooth variants, stacked-100 line: 3. Phase B: secondary axis and combination charts: 2. Phase C: 3D and cone/cylinder/pyramid: 2. Phase D: extended charts (chartEx part: treemap, sunburst, waterfall, funnel, histogram, box & whisker): 4 |
| Chart options | Trendlines, error bars, data labels from cells, manual label layout, high-low/drop/up-down bars, axis type change, 100+ preset styles, `.crtx` templates, shapes inside charts | Title, axis titles, legend, data labels (4 modes), series and marker colours, value-axis format | Medium, Business | Trendlines, error bars, axis min/max/major unit, gridlines, number formats on both axes, data-label position, gap width/overlap, smooth lines, per-point colours: 2. Preset styles: 2. Templates: 1 |
| Sparklines | Line, column, win/loss with full options | No | Small, Business | 1.5 |
| Pictures | 11 formats incl. svg, webp, emf/wmf, tiff, bmp, ico; linked pictures; effects; in-cell pictures; `IMAGE()` | PNG, JPEG, GIF | Small, Business | BMP, TIFF, WebP, SVG, ICO (dimension parsing only; the bytes are stored as-is): 1. In-cell pictures (`vm`/rich value parts): 2 |
| Shapes, text boxes (187 shape types), grouping, effects, adjustment points | Yes | No | Medium, occasional | Text boxes and the 20 common shapes with fill, line, text: 2. Grouping and the full set: 2 |
| Form controls (buttons, check boxes, drop-downs, ...) | Yes | No | Niche | Defer |
| OLE objects, chart/shape image export | Yes / 9-preview | No | Niche | Non-goal |
| Drawing positioning (two-cell anchors, pixel sizes, copy drawings) | Yes | Cell anchor with pixel size | Small | Two-cell anchor and absolute offsets: 0.5 |

### 3.8 Formulas

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Write formulas, shared/array formulas, R1C1, `_xlfn.` prefixes added automatically | Yes | Write any formula text; shared and array formulas preserved in edit mode; no R1C1 | Small | Add `_xlfn.` prefixes for the 2010+ function list automatically; `WriteFormulaR1C1`: 0.5 |
| Cached results written | From the engine | Explicit value, or the opt-in evaluator (`SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `ROUND`, `ABS`, arithmetic, same-sheet references) | Part of the engine gap | — |
| Calculation engine: 478 functions, dependency chain, dynamic arrays, LAMBDA, external links, custom functions, cancellation, precision options | Yes | No. Excel calculates on open (`fullCalcOnLoad`); `ExcelReader` returns cached results | **Largest gap, Business when reading calculated totals from files you did not write** | Phase 1, the evaluator becomes a workbook engine for files SheetStream wrote or read: parser with precedence, ranges, cross-sheet references, names, dependency ordering, and the 120 most used functions (logical, text, date, math, lookup incl. `IF`, `IFS`, `VLOOKUP`, `XLOOKUP`, `INDEX`/`MATCH`, `SUMIF(S)`, `COUNTIF(S)`, `TEXT`, `DATE`, `EOMONTH`, `ROUND*`, `CONCAT`, `LEFT/RIGHT/MID`, `IFERROR`): 12. Phase 2, statistical/financial/engineering to about 300 functions: 10. Phase 3, dynamic arrays and spill, `LET`, `LAMBDA`: 10. Pivot calculation on top: 5 |
| `Calculate()` on read (`ExcelReader`) | Yes | No | Same | With the engine: a `ReadCalculated` mode that recalculates dirty cells |

### 3.9 Data validation, filtering, sorting

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Validation types | Whole number, decimal, list, date, time, text length, custom, any | All but **time** | Small | 0.3 |
| Validation options | Prompt, error (stop/warning/info), `AllowBlank`, `HideDropDown`, cross-sheet lists | Prompt, error styles, inline and range lists | Small | `AllowBlank`, `HideDropDown`, `ShowInputMessage` toggles: 0.3 |
| AutoFilter with criteria, `ApplyFilter` | Yes | Filter buttons | Medium | See 3.6 |
| Range sort, sort by colour, custom lists | Yes | No | Small | `sortState` written; data sorted by the caller: 0.5 |
| Validation enforcement | Not enforced (documented) | Not enforced | None | — |

### 3.10 Printing

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Page setup (orientation, paper, scale, fit, margins, centre, gridlines, headings, black & white, draft, page order) | Yes | All but headings, black & white, draft, page order | Small | 0.3 |
| Print area and titles | Yes | Yes | None | — |
| Manual page breaks | Yes | No | Small, Business (invoices) | Row and column breaks: 0.5 |
| Headers and footers: codes, pictures, odd/even, first page | Yes | Text with codes on all pages | Small | Odd/even/first variants: 0.3; pictures: 1 |

### 3.11 Reading and import

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| Typed values, dates (1904 too), errors, durations | Yes | Yes | None | — |
| Formatted display text (`Text`) | Yes, with culture handling | No | Small, Business (import "what the user saw") | Number-format renderer for built-in and common custom codes: 2 |
| Streaming, async, cancellation | Load whole package; `LoadAsync` | Yes | **SheetStream better** | — |
| `.xls` (BIFF8) | Not supported | Not supported | None versus EPPlus; **Business** versus ExcelDataReader | BIFF8 reader (workbook globals, SST, BOF/BOUNDSHEET, RK/NUMBER/LABELSST/FORMULA cached values, number formats, 1904): 5. Makes SheetStream a full replacement for ExcelDataReader |
| CSV | Yes | No | Small, Business | See 3.2 |
| Read definitions: tables, validations, conditional formats, charts, pivots, comments, defined names, drawings | Yes | Merged ranges, hyperlinks; everything else is preserved only through template/edit copying | Medium | Read tables (name, range, columns), defined names, comments, validations and conditional formats as objects: 2. Charts and pivots as objects: defer |
| Lenient header matching, conversion-failure reporting | Mapping options, `ConversionFailureStrategy` | Case-insensitive match; row and column named in the exception | Small, Business | See 3.5 |
| Workbook properties | Yes | Write only | Small | Read `docProps`: 0.2 |

### 3.12 Performance and memory

| Measure | EPPlus 8 | SheetStream | Verdict |
|---|---|---|---|
| Write 100,000 × 10 formatted | EPPlus 4.5.3.3: 2.3 s, 390 MB allocated, 184 MB peak | 1.0 s (0.53 s fastest zip), 27 MB allocated, 30 MB peak | SheetStream better |
| Write 100,000 × 200 (EPPlus' own benchmark) | 29.1 s on an i9, whole workbook in memory | Not measured on that shape; 1,000 × 10,000 (same cell count) runs with a 160 MB peak | SheetStream better on memory; add the 100,000 × 200 shape to the suite for a direct number (0.3 days) |
| Read 100,000 × 10 | EPPlus 4.5.3.3 loads the package (hundreds of MB) | 2.7 MB above baseline, streaming | SheetStream better |
| Large workbooks (millions of cells) | Memory-bound | Flat memory for writing and reading; editing is bounded by the edit limit | SheetStream better for writing and reading; EPPlus better for editing |

### 3.13 Tooling and support

| Feature | EPPlus 8 | SheetStream | Gap | Plan (days) |
|---|---|---|---|---|
| API reference per version, wiki (132 pages), FAQ | Yes | DocFX site with 7 guides and the API reference | Medium | Grow the guides to one per feature area (styling, charts, tables, validation, printing, reading): 2 |
| Sample repositories (C#, VB, web, Docker) | Yes | Showcase console sample, web sample, DotNetFiddle | Small | A samples folder with one small program per feature area: 1.5 |
| Analyzer and source generator | No | Yes | **SheetStream better** | — |
| Commercial support, release cadence | Yes; minor every 1 to 3 months | GitHub; releases as features land | By design | — |
| Compatibility proof against EPPlus | — | Benchmarks against EPPlus 4.5.3.3 | Small | A test that builds the EPPlus 4.5.3.3 samples with both libraries and compares cell values and styles: 1.5 |

## 4. What not to chase

These EPPlus features would cost more than they are worth for SheetStream's users, or conflict with its design rules (zero dependencies, flat memory). They are documented as non-goals rather than left as silent gaps.

- **Sensitivity labels (Microsoft Information Protection).** Needs the MIP SDK, Purview and an Entra app registration: a dependency and an operational footprint SheetStream should not carry.
- **OLE embedded objects, form controls, external connections and query tables.** Legacy or Power-Query-specific; no business export needs them.
- **The full 478-function engine before the common 120.** The functions that matter are the ones that appear in reports and templates; the long tail (engineering, database, legacy statistical) can wait for a request.
- **Region map, surface, cone/cylinder/pyramid charts.** Excel builds the map data; the rest are novelty 3D variants. Do them last, if ever.
- **VBA module creation and signing.** Preserving macros in `.xlsm` is worth doing; generating VBA is not.
- **PDF and SVG export.** EPPlus 9 preview only; a different product. If reports need PDF, that is a separate component.

## 5. Roadmap to parity

Ordered by business value divided by cost, with the first production migration's priorities (LogisticaHQ) first. Each phase ends with a release.

| Phase | Content | Days | Release |
|---|---|---|---|
| **0 (in progress)** | AutoFit with font width table and measurer hook; benchmark tables; v0.4.0 | 1 | 0.4.0 |
| **1. Replace ExcelDataReader** | CSV read and write; BIFF8 `.xls` reader; lenient header matching, `[DisplayName]` already, conversion-failure strategy; range and table scope for `Read<T>`/`ReadDataTable`; encrypted `.xlsx` reading; `.xlsm` round-trip test; `GetValue<T>` | 11 | 0.5.0 |
| **2. Editing parity for everyday files** | Insert/delete rows and columns with reference shifting; copy ranges and sheets; larger in-place edits by keeping untouched rows raw; named styles; gradient fills; theme colours; remaining conditional-format rules; sheet reorder; protected ranges; page breaks; header/footer variants and pictures; time validation and validation flags | 13 | 0.6.0 |
| **3. Tables, pivots, sparklines** | Table totals and calculated columns; filter criteria and `sortState`; pivot cache with data; grouping; calculated fields; Show Values As; pivot filters and sorting; slicers (table and pivot); sparklines; pivot charts (after phase 4A) | 15 | 0.7.0 |
| **4. Charts** | A: 2D type parity; B: secondary axis and combination charts; chart options (trendlines, error bars, axis control, label layout); more image formats; text boxes and common shapes; two-cell anchors | 11 | 0.8.0 |
| **5. Export and reading definitions** | JSON and CSV export; formatted `Text`; read tables, names, comments, validations, conditional formats; HTML export; compatibility test against EPPlus 4.5.3.3 samples; samples folder and per-area guides | 11 | 0.9.0 |
| **6. Formula engine** | Phase 1 engine (120 functions, cross-sheet, names, dependency order) for write-time results and `ReadCalculated`; then statistical/financial breadth; then dynamic arrays and `LAMBDA`; pivot calculation | 12 + 10 + 10 + 5 | 1.0.0 |
| **7. Long tail** | 3D and extended chart types, chart presets and templates, threaded comments, digital signatures (net8.0), form controls, in-cell pictures | 15 | after 1.0 |

Phases 1 to 5 reach business parity in about 60 agent-days; phase 6 is what makes "better than EPPlus in every aspect" true for formulas, and it is the one item where EPPlus' head start is real. Every phase keeps the three rules that make SheetStream better today: no dependencies, flat memory on the streaming path, and verified output (schema validation plus real Excel through COM for anything that Excel is picky about, as pivots were).

## 6. Sources

EPPlus 8: feature page (epplussoftware.com/en/Developers/Features), EPPlus 8 and 7 release pages, licence overview and FAQ, NuGet package metadata (8.7.1), the EPPlus.csproj on `develop8`, the GitHub wiki (Getting Started, Configuration, Breaking Changes in 7 and 8, and the per-feature pages), the 8.7 API reference, the GitHub releases list and the 9.0.0-preview notes. SheetStream: this repository at `2bcb8fa`, its README, CHANGELOG, tests and `benchmarks/README.md`. Items EPPlus Software does not document (thread safety, `.xlsb`, formal AOT support, pivot `RefreshOnLoad`, colour filters) are treated as absent.
