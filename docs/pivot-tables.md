# Pivot tables

A pivot table summarises a range, or an Excel table, written earlier in the same workbook. SheetStream writes the pivot's layout and asks Excel to compute it when the file opens, so the result is a real, interactive pivot: users can drag fields, filter, drill and refresh exactly as if they had built it themselves.

```csharp
using SheetStream;

using var package = ExcelPackage.Create("sales.xlsx");
var wb = package.Workbook;

var data = wb.Worksheets.Add("Data");
var loaded = data.LoadFromCollection(orders, new LoadOptions { TableName = "Orders" });

var report = wb.Worksheets.Add("Pivot");
var pivot = report.AddPivotTable("SalesByRegion", loaded.Table!, row: 3, column: 1);
pivot.AddRowField("Region");
pivot.AddColumnField("Category");
pivot.AddDataField("Amount", PivotFunction.Sum, numberFormat: "$#,##0.00");
pivot.AddDataField("OrderId", PivotFunction.Count, displayName: "Orders");

package.Save();
```

## Two ways to name the source

| Overload | Source | Field names come from |
|---|---|---|
| `AddPivotTable(name, excelTable, row, column)` | An `ExcelTable` from `AddTable` or `LoadOptions.TableName` (`LoadResult.Table`) | The table's columns. Excel re-reads the whole table on each refresh, so the pivot follows the table when rows are added later in Excel. |
| `AddPivotTable(name, sourceSheetName, sourceRange, row, column, headers)` | Any A1 range including its header row, e.g. `"A1:F500"` | The `headers` you pass. `LoadResult.Headers` and `LoadResult.Range` give you both for data written with a `LoadFrom*` call. |

The header texts are needed because the source sheet has already been streamed into the file: SheetStream cannot read it back. Field names are matched case-insensitively; an unknown name throws an `ArgumentException` that lists the available headers.

The source sheet may be the pivot's own sheet (a pivot next to its data) or any sheet added earlier. It must exist in the workbook when `AddPivotTable` is called.

## Placing fields

| Method | Excel area | Notes |
|---|---|---|
| `AddRowField(header)` | Rows | Several row fields nest in the order they are added. |
| `AddColumnField(header)` | Columns | |
| `AddFilterField(header)` | Filters (report filter / page field) | Each filter takes one row above the pivot body, followed by one blank row, as Excel lays it out. `row` is the first filter cell. |
| `AddDataField(header, function, displayName, numberFormat)` | Values | `PivotFunction`: Sum (default), Count, Average, Max, Min, Product, CountNumbers, StdDev, StdDevP, Var, VarP. The default caption is `"Sum of Amount"` style. With more than one data field Excel shows them side by side under a "Values" header. |

A data field's caption must differ from every source header. Excel itself refuses `displayName: "Amount"` when a column is called Amount ("PivotTable field name already exists"), so SheetStream rejects it up front.

Options on `PivotTable`:

- `StyleName`: a built-in pivot style, default `"PivotStyleMedium9"`; also `ShowRowStripes` / `ShowColumnStripes`.
- `ShowRowGrandTotals` (the Grand Total column at the right) and `ShowColumnGrandTotals` (the Grand Total row at the bottom), both default `true`.
- `Compact`: `true` (default) stacks all row fields in one indented column, as Excel creates pivots; `false` gives outline form with one column per row field.

Several pivots can share a sheet, sit on different sheets or point at the same source; each gets its own cache. Pivot names must be unique within the workbook.

## How it works: refresh on open

Streaming means the data is gone by the time the pivot is written, so SheetStream does not build the pivot cache (the list of distinct items per field and the record table) that Excel normally stores next to a pivot. Instead the file carries:

- `xl/pivotCache/pivotCacheDefinitionN.xml` with the source reference, one `cacheField` per header, no items, `recordCount="0"`, `saveData="0"` and `refreshOnLoad="1"`;
- `xl/pivotCache/pivotCacheRecordsN.xml` with zero records (Excel expects the part to exist even when it is empty);
- `xl/pivotTables/pivotTableN.xml` with the fields, their areas, the style and a `location` sized for one item per axis.

`refreshOnLoad` makes Excel read the source cells and compute every total the moment the workbook opens, exactly as if the user had clicked Refresh All, and it then lays the pivot out for the real number of items. Opening and saving the file in Excel stores a full cache like any other pivot.

Version stamps are written as Excel 2016 writes them (`createdVersion`/`refreshedVersion` 6, `minRefreshableVersion` 3): Excel treats a cache whose refreshed version is older than the pivot's as stale and some builds then refuse to refresh it.

`tests/excel-check-pivot.ps1` opens a workbook through Excel COM automation, checks that no part was dropped by a repair, calls `RefreshTable()` on every pivot and prints the refreshed cells; it is how the shipped pivots were verified.

## Limits

- **No cached data.** A viewer that does not run the pivot engine (web previews, some mobile apps, older LibreOffice builds) shows the header cells only, until the file has been opened and saved in Excel. If the report must render everywhere, write the summary as plain cells (`GroupBy` plus `WriteCell`) or formulas (`SUMIFS` on the table) instead.
- Supported layouts: row, column and filter fields, data fields with the eleven standard functions, number formats, grand totals, compact or outline form, built-in styles.
- Not supported: calculated fields and items, grouping (dates, numbers), slicers and timelines, custom sort orders, "show values as" (percent of total, running total), conditional formats inside the pivot, OLAP sources, external connections.
- Pivots cannot be added to template or edited workbooks (`OpenTemplate`, `Edit`), and existing pivots are not read or modified.
- A pivot whose source range has only a header row is refused (`ArgumentException`), as Excel cannot build a cache from it.
