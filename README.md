# 🚀 SheetStream

A lightning-fast, zero-dependency, **100% MIT-licensed** .NET library for creating and reading Excel (`.xlsx`) files.

Stop paying thousands of dollars in commercial developer subscriptions just to write a spreadsheet. **SheetStream** uses a highly optimized, streaming XML engine to build massive Excel files with ultra-low memory allocation and absolute legal freedom.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![NuGet](https://img.shields.io/nuget/v/SheetStream.svg)](https://www.nuget.org/packages/SheetStream)
[![Build](https://github.com/P4Software/SheetStream/actions/workflows/build.yml/badge.svg)](https://github.com/P4Software/SheetStream/actions/workflows/build.yml)

## ✨ Why SheetStream?

* **⚖️ Zero Financial & Legal Friction:** 100% MIT Licensed. Free for commercial use, personal use, startups, and massive enterprises alike. No license keys, no tier audits, and no required subscriptions.
* **📦 Zero Dependencies:** Built entirely on native .NET `System.IO.Compression` and `System.Xml` APIs. Targets `netstandard2.0` and `net8.0`, so it runs on .NET Framework 4.7.2+ as well as modern .NET.
* **⚡ High-Performance Streaming:** Rows go straight through an `XmlWriter` into the zip archive. 10 million cells write in about 5 seconds with a 41 MB peak working set, whether the sheet is long or wide.
* **🧰 The features real reports use:** tables, charts, pictures, conditional formatting, data validation, hyperlinks, comments, rich text, protection, print setup, outlining, AutoFit, and a streaming reader for imports.
* **✅ Verified ECMA-376 output:** Every feature is validated against the OpenXML schema with the official Open XML SDK validator, and the showcase workbook is opened in real Excel to confirm nothing triggers a repair.

## 📦 Installation

```bash
dotnet add package SheetStream
```

```powershell
Install-Package SheetStream
```

## 💻 Quick Start

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
            .WriteCell($"Enterprise SKU-{i}")          // Shared string (auto-deduplicated)
            .WriteCell(1500.50 * i, currencyStyle);
    }

    sheet.Row(102).WriteCell("Total", headerStyle).WriteEmpty(headerStyle).WriteFormula("SUM(C2:C101)", currencyStyle);
    package.Save(); // Also happens automatically when the package is disposed
}
```

Need bytes for a web download, like EPPlus' `GetAsByteArray()`?

```csharp
byte[] bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Export").LoadFromCollection(orders));
```

## 🧭 Feature tour

### Load data in one call

```csharp
var result = sheet.LoadFromCollection(orders, new LoadOptions
{
    HeaderStyle = headerStyle,
    TableName = "Orders",              // Becomes an Excel table with filter buttons and banding
    TableStyle = "TableStyleMedium2",
});
sheet.LoadFromDataTable(dataTable);
sheet.LoadFromDataReader(sqlCommand.ExecuteReader()); // Streams a large query row by row
```

Control columns with attributes: `[SheetColumn("Invoice #", Order = 0, NumberFormat = "$#,##0.00")]`, `[SheetColumn(Ignore = true)]`, or `[DisplayName("...")]`. `IDictionary<string, object>` rows (ExpandoObject, Dapper) work too.

### Tables, conditional formatting and validation

```csharp
var table = sheet.AddTable("Sales", "A1:D101");      // Add before writing the header row
table.StyleName = "TableStyleMedium9";

var bad = workbook.Styles.CreateDifferential(o => { o.Font.Color = "#9C0006"; o.Fill.BackgroundColor = "#FFC7CE"; });
sheet.AddConditionalFormatting("D2:D101")
     .CellValue(ConditionalOperator.LessThan, "0", bad)
     .DataBar("#638EC6");
sheet.AddConditionalFormatting("E2:E101").IconSet(IconSetType.ThreeTrafficLights);

sheet.AddDataValidation("C2:C101").List("Open", "Closed", "On hold").WithError("Invalid", "Pick from the list.");
sheet.AddDataValidation("F2:F101").WholeNumber(ValidationOperator.Between, 1, 100);
```

Conditional formatting covers cell value, expression, contains text, duplicates, unique, top / bottom N, above / below average, two- and three-colour scales, data bars and icon sets. Validation covers lists (inline or from a range), whole numbers, decimals, dates, text length and custom formulas.

### Charts and pictures

```csharp
var chart = sheet.AddChart(ChartType.Column, "H2:P20");
chart.Title = "Revenue by month";
chart.YAxisTitle = "USD";
chart.ValueAxisNumberFormat = "$#,##0";
chart.AddSeries(values: "C2:C13", categories: "A2:A13", name: "=C1", color: "#4F81BD");

sheet.AddPicture("logo.png", row: 1, column: 8, widthPixels: 160);
```

Chart types: column, bar (clustered, stacked, 100%), line, line with markers, stacked line, area, stacked area, pie, doughnut, scatter and scatter with lines. Pictures: PNG, JPEG and GIF.

### Links, comments, rich text and formulas

```csharp
sheet.Row()
     .WriteHyperlink("Open the portal", "https://example.com/orders/42")
     .WriteInternalLink("Go to summary", "'Summary'!A1")
     .WriteRichText("Status: ", new RichText("LATE") { Bold = true, Color = "#C00000" })
     .WriteFormula("B2*TaxRate", cachedValue: 7.0);   // Cached result for viewers that do not calculate
sheet.AddComment(2, 3, "Checked by finance", author: "Glenn");
workbook.DefinedNames.Add("TaxRate", "'Settings'!B2");
```

### Layout, printing and protection

```csharp
sheet.AutoFitColumns();                       // Widths sized to content at save time
sheet.TabColor = "#C00000";
sheet.ShowGridLines = false;
sheet.Columns(5, 8).SetOutlineLevel(1);       // Collapsible column group
sheet.Row(10, new RowOptions { OutlineLevel = 1, Hidden = true });

sheet.PageSetup.Orientation = PageOrientation.Landscape;
sheet.PageSetup.FitToWidth = 1;
sheet.PageSetup.PrintTitleRows = (1, 1);
sheet.PageSetup.Footer = "Page &P of &N";

var input = workbook.Styles.Create(o => o.Protection.Locked = false);
sheet.Protect("password", o => o.AllowSort = true);   // Salted SHA-512, like Excel 2013+
workbook.Protect("password");                         // Lock the sheet structure
sheet.Visibility = SheetVisibility.Hidden;
```

### Random access, EPPlus style

```csharp
sheet.Cells[3, 2].Value = 42;          // Any order
sheet.Cells["B4"].Formula = "SUM(B1:B3)";
sheet.Cells["B4"].Style = boldStyle;
```

Random-access mode keeps the sheet's cells in memory until it is saved, which is convenient for small, irregular layouts. Use `Row()` streaming for big exports. A sheet uses one mode or the other.

### Reading workbooks

```csharp
using SheetStream.Reading;

using var reader = ExcelReader.Open("import.xlsx");
foreach (var row in reader.ReadRows("Products"))           // Streams; constant memory
    Console.WriteLine($"{row.GetString(0)}: {row.GetDouble(2)}");

List<Product> products = reader.Read<Product>("Products").ToList();   // Header row maps to properties
DataTable table = reader.ReadDataTable("Products");
```

Dates come back as `DateTime` (built-in and custom date formats, 1904 date system included), elapsed times as `TimeSpan`, errors as `ExcelErrorValue`, and formula cells as their cached result.

## 📚 API overview

| Area | What you get |
|---|---|
| `ExcelPackage` | `Create(path)`, `new ExcelPackage(stream)`, `CreateBytes(build)`, `Save()`, `Dispose()` |
| `Workbook` | `Worksheets.Add`, `Styles`, `DefinedNames`, `Properties`, `Protect` |
| `Worksheet` | `Row()`, `Cells[r, c]`, `Column` / `Columns`, `AutoFitColumns`, `FreezePanes`, `AutoFilter`, `MergeCells`, `AddTable`, `AddChart`, `AddPicture`, `AddConditionalFormatting`, `AddDataValidation`, `AddHyperlink`, `AddComment`, `Protect`, `PageSetup`, `TabColor`, `Visibility`, `LoadFromCollection` / `DataTable` / `DataReader` |
| `RowWriter` | `WriteCell` (string, numbers, bool, dates, TimeSpan, object), `WriteCells`, `WriteFormula` (with or without cached value), `WriteRichText`, `WriteHyperlink`, `WriteInternalLink`, `WriteInlineString`, `WriteEmpty`, `Skip` |
| `StyleSheet` | `Create` (font, fill, border, number format, alignment, wrap, indent, rotation, shrink, protection), `CreateDifferential`, `Hyperlink` |
| `ExcelReader` | `Open`, `SheetNames`, `ReadRows`, `Read<T>`, `ReadDataTable` |

## ⚖️ SheetStream vs EPPlus 8

| | EPPlus 8 | SheetStream |
|---|---|---|
| License | Commercial key for business use | MIT |
| Dependencies | Several packages | None |
| Write cells, styles, formats, widths, freeze, filters, merges | ✅ | ✅ |
| Tables, charts, pictures, conditional formatting, validation | ✅ | ✅ (common chart types) |
| Hyperlinks, comments, rich text, defined names, protection, print setup, outlining | ✅ | ✅ |
| AutoFit columns | ✅ font metrics | ✅ character-count estimate |
| Load from collection / DataTable / DataReader | ✅ | ✅ |
| Random-access `Cells[r, c]` | ✅ | ✅ (per sheet, in memory) |
| Read workbooks | ✅ | ✅ streaming, read-only |
| Large exports | Whole workbook in memory | Streaming, flat memory |
| Edit and re-save an existing workbook | ✅ | ❌ |
| Formula calculation engine | ✅ | ❌ Excel calculates on open; cached values can be supplied |
| Pivot tables | ✅ | ❌ |
| Sparklines, OLE objects, digital signatures, sensitivity labels, VBA, encryption | ✅ | ❌ |

### Migrating from EPPlus

| EPPlus | SheetStream |
|---|---|
| `new ExcelPackage()` + `package.SaveAs(file)` | `ExcelPackage.Create(path)` + `package.Save()` |
| `package.GetAsByteArray()` | `ExcelPackage.CreateBytes(wb => ...)` |
| `ws.Cells[r, c].Value = x` | `ws.Cells[r, c].Value = x`, or stream with `ws.Row(r).WriteCell(x)` |
| `cell.Style.Font.Bold = true` | `Styles.Create(o => o.Font.Bold = true)` once, pass the style to `WriteCell` |
| `cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79,129,189))` | `o.Fill.BackgroundColor = Colors.Rgb(79, 129, 189)` |
| `cell.Style.Numberformat.Format = "#,##0.00"` | `o.NumberFormat = NumberFormats.Decimal2` |
| `ws.Column(c).Width = 15` | `ws.Column(c).SetWidth(15)` |
| `ws.Cells[...].AutoFitColumns()` | `ws.AutoFitColumns()` before writing |
| `ws.Cells["A1"].LoadFromCollection(list, true, TableStyles.Medium2)` | `ws.LoadFromCollection(list, new LoadOptions { TableName = "T" })` |
| `ws.Tables.Add(range, "Sales")` | `ws.AddTable("Sales", "A1:D101")` |
| `ws.Drawings.AddChart("c", eChartType.ColumnClustered)` | `ws.AddChart(ChartType.Column, "H2:P20")` |
| `ws.Cells["A1"].Hyperlink = new Uri(url)` | `row.WriteHyperlink(text, url)` |
| `new ExcelPackage(file).Workbook.Worksheets[0].Cells[r, c].Value` | `ExcelReader.Open(file).ReadRows(0)` |

### Streaming rules

SheetStream writes directly into the zip archive, which is what keeps memory flat. Each rule is enforced with a clear exception:

1. **One sheet at a time.** Adding a worksheet seals the previous one.
2. **Rows ascend.** `Row(n)` must be the current row or a later one. Gaps are fine. Use `Cells[r, c]` on a sheet that needs random access.
3. **Layout first.** Column widths, frozen panes, tab colour, gridlines and fit-to-page live before the data in the file, so set them before the first row. `AutoFitColumns()` and `Cells[...]` defer the header, which lifts this rule for that sheet.
4. **Dispose saves.** Leaving a `using` block without calling `Save()` still finishes the file. A save that fails is not retried by `Dispose`, so the original exception is the one you see.

## 🛠️ Architecture

1. **Package (`ExcelPackage`):** owns the `ZipArchive`, streams each worksheet part as it is written, and on save emits the content types, relationships, workbook, styles, shared strings and document properties.
2. **Streaming writer (`Worksheet` / `RowWriter`):** rows go through an `XmlWriter` into the compressed stream. Column letters are cached and the row number is formatted once per row, so writing a cell allocates nothing. Tables, drawings, charts and comments are separate parts written right after their sheet.
3. **Shared strings (`SharedStringTable`):** de-duplicates text across all sheets, as Excel does.
4. **Styles (`StyleSheet`):** de-duplicates fonts, fills, borders, number formats, cell formats and differential formats, written in the exact element order Excel demands.
5. **Reader (`ExcelReader`):** an `XmlReader` pass over each sheet that resolves shared strings and date styles, with DTD processing disabled for untrusted files.

## 🧪 Building and testing

```bash
dotnet build
dotnet test
dotnet run --project samples/SheetStream.Sample                         # FinancialReport.xlsx
dotnet run --project samples/SheetStream.Sample -- showcase              # Showcase.xlsx: every feature
dotnet run -c Release --project samples/SheetStream.Sample -- 1000000    # 1M-row stress run
dotnet run -c Release --project samples/SheetStream.Sample -- 5000 2000  # 10M-cell wide stress run
```

## 🚢 Releasing

Bump `<Version>` in `src/SheetStream/SheetStream.csproj`, add a `CHANGELOG.md` entry, commit, then:

```bash
git tag v0.2.0
git push --tags
```

The `Release` workflow refuses to run if the tag and the csproj version disagree. It publishes through nuget.org Trusted Publishing, so no API key is stored in the repository.

## 🤝 Contributing

We welcome the .NET community to help build the best free spreadsheet engine out there! See [CONTRIBUTING.md](CONTRIBUTING.md).

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.
