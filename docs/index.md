# SheetStream

A zero-dependency, MIT-licensed .NET library for creating and reading Excel (`.xlsx`) files. Rows stream straight into the file, so 10 million cells write in seconds with flat memory.

```bash
dotnet add package SheetStream
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

## Formulas and cached results

`row.WriteFormula("SUM(C2:C101)")` writes the formula and asks Excel to calculate on open. Viewers that do not calculate (file previews, Quick Look, many mobile and web viewers, `ExcelReader`) show such a cell blank unless the file also carries the result. Turn on `sheet.CacheFormulaResults()` before the first row and SheetStream computes the result itself when the formula is simple enough: numbers, references to cells already written on the same sheet, ranges, `+ - * /`, `%`, parentheses and `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`, `ROUND`, `ABS`. Anything else (other functions, other sheets, rows not yet written, a division by zero) is written exactly as before, without a value; a wrong value is never written. The sheet keeps about 8 bytes per numeric cell to do this, which is why it is opt-in. When you already know the result, `WriteFormula(formula, cachedValue)` stores it directly.

## Where to go next

- [Migrating from EPPlus](migrating-from-epplus.md): a real exporter converted line by line.
- [Filling templates](templates.md): design in Excel, fill from code.
- [API reference](api/SheetStream.yml): every public type, generated from the source.
- [Benchmarks](https://github.com/P4Software/SheetStream/tree/main/benchmarks): reproducible comparisons with EPPlus, ClosedXML, MiniExcel and NPOI.

## Design in one paragraph

SheetStream writes each worksheet through an `XmlWriter` directly into the zip archive, so it never builds the workbook in memory. That buys flat memory and speed, and costs random access: rows are written in ascending order, one sheet at a time, and layout that Excel stores before the data (column widths, frozen panes) is declared first. Sheets that need random access can use `sheet.Cells[row, column]`, which buffers that one sheet. Everything the library writes is validated against the ECMA-376 schema in CI and opened in real Excel before each release.
