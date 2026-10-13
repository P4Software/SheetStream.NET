# 🚀 SheetStream

A lightning-fast, zero-dependency, **100% MIT-licensed** .NET library for generating high-performance Excel (`.xlsx`) files.

Stop paying thousands of dollars in commercial developer subscriptions just to write a spreadsheet. **SheetStream** uses a highly optimized, streaming XML engine to build massive Excel files with ultra-low memory allocation and absolute legal freedom.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![NuGet](https://img.shields.io/nuget/v/SheetStream.svg)](https://www.nuget.org/packages/SheetStream)
[![Build](https://github.com/P4Software/SheetStream/actions/workflows/build.yml/badge.svg)](https://github.com/P4Software/SheetStream/actions/workflows/build.yml)

## ✨ Why SheetStream?

* **⚖️ Zero Financial & Legal Friction:** 100% MIT Licensed. Free for commercial use, personal use, startups, and massive enterprises alike. No sudden license changes, no tier audits, and no required subscriptions.
* **📦 Zero Dependencies:** Built entirely on native .NET `System.IO.Compression` and `System.Xml` APIs. No bloated third-party wrappers. Targets `netstandard2.0` and `net8.0`, so it runs on .NET Framework 4.7.2+ as well as modern .NET.
* **⚡ High-Performance Streaming:** Unlike legacy libraries that load entire document object models (DOM) into memory, SheetStream writes XML sequentially straight into the zip archive. Generate millions of rows without triggering Out-Of-Memory exceptions or crippling your RAM.
* **🛠️ Modern Fluent API:** Write clean, readable, and chainable C# code to format your workbooks.
* **✅ Verified ECMA-376 output:** The test suite validates every generated file against the OpenXML schema with the official Open XML SDK validator, so the files open cleanly in Excel, Google Sheets, and LibreOffice.

## 📦 Installation

Install via the .NET CLI:

```bash
dotnet add package SheetStream
```

Or via the NuGet Package Manager Console:

```powershell
Install-Package SheetStream
```

## 💻 Quick Start

Create a styled 3-column, 100-row spreadsheet in seconds:

```csharp
using System;
using SheetStream;
using SheetStream.Styles;

using (var package = ExcelPackage.Create("FinancialReport.xlsx"))
{
    var workbook = package.Workbook;
    var sheet = workbook.Worksheets.Add("Q1 Report");

    // Configure columns (before the first row: the sheet is streamed forward-only)
    sheet.Column(1).SetWidth(10);
    sheet.Column(2).SetWidth(28);
    sheet.Column(3).SetWidth(18);
    sheet.FreezePanes(rows: 1);

    // Styles are registered once and de-duplicated across the workbook
    var headerStyle = workbook.Styles.Create(options =>
    {
        options.Font.Bold = true;
        options.Font.Color = "#FFFFFF";
        options.Fill.BackgroundColor = "#4F81BD";
        options.Border.Style = BorderStyle.Thin;
    });
    var currencyStyle = workbook.Styles.Create(options => options.NumberFormat = "$#,##0.00");

    // Header row
    sheet.Row(1)
        .WriteCell("ID", headerStyle)
        .WriteCell("Product Name", headerStyle)
        .WriteCell("Revenue (USD)", headerStyle);

    // High-performance data streaming
    for (int i = 1; i <= 100; i++)
    {
        sheet.Row(i + 1)
            .WriteCell(i)                               // Number
            .WriteCell($"Enterprise SKU-{i}")           // Shared string (auto-deduplicated)
            .WriteCell(1500.50 * i, currencyStyle);     // Number with format
    }

    // Totals row with a live formula
    sheet.Row(102).WriteCell("Total", headerStyle).WriteEmpty(headerStyle).WriteFormula("SUM(C2:C101)", currencyStyle);

    package.Save(); // Also happens automatically when the package is disposed
}
```

Need the file as bytes for a web download? Mirror EPPlus' `GetAsByteArray()`:

```csharp
byte[] bytes = ExcelPackage.CreateBytes(workbook =>
{
    var sheet = workbook.Worksheets.Add("Export");
    sheet.Row().WriteCells("Id", "Name", "When");
    sheet.Row().WriteCells(1, "Alpha", DateTime.Today);
});
```

## 📚 API Overview

| Area | What you get |
|---|---|
| `ExcelPackage` | `Create(path)`, `new ExcelPackage(stream)`, `CreateBytes(build)`, `Save()`, `Dispose()` |
| `Workbook` | `Worksheets.Add(name)`, `Styles.Create(...)`, `SharedStrings` counters |
| `Worksheet` | `Column(i).SetWidth(w).Hide().SetStyle(s)`, `FreezePanes(rows, cols)`, `AutoFilter(row, c1, c2)`, `MergeCells(...)`, `Row()`, `Row(index)`, `Row(index, style, height)` |
| `RowWriter` | `WriteCell(...)` for `string`, `double`, `decimal`, `int`, `long`, `bool`, `DateTime`, `DateTimeOffset`, `TimeSpan` and `object`; `WriteCells(params)`, `WriteInlineString`, `WriteFormula`, `WriteEmpty`, `Skip` |
| `StyleOptions` | `Font` (bold, italic, underline, strike, size, name, color), `Fill.BackgroundColor`, `Border` (per edge + color), `NumberFormat`, `HorizontalAlignment`, `VerticalAlignment`, `WrapText` |
| `PackageOptions` | `CompressionLevel`, `UseSharedStrings` (switch to inline strings for data sets dominated by unique text) |

### Streaming rules

SheetStream writes directly into the zip archive, which is what keeps memory flat. That implies a few rules, each enforced with a clear exception:

1. **One sheet at a time.** Adding a worksheet seals the previous one.
2. **Rows ascend.** `Row(n)` must be the current row or a later one. Gaps are fine.
3. **Layout first.** Column widths and frozen panes are declared before the first row of the sheet, because the OpenXML schema places them before the sheet data. `AutoFilter` and `MergeCells` can be called any time before the sheet is sealed.

### Migrating from EPPlus

| EPPlus | SheetStream |
|---|---|
| `new ExcelPackage()` + `package.SaveAs(file)` | `ExcelPackage.Create(path)` + `package.Save()` |
| `package.GetAsByteArray()` | `ExcelPackage.CreateBytes(wb => ...)` |
| `ws.Cells[r, c].Value = x` | `ws.Row(r).WriteCell(x)` (cells left to right) |
| `cell.Style.Font.Bold = true` | `Styles.Create(o => o.Font.Bold = true)` once, pass the style to `WriteCell` |
| `cell.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79,129,189))` | `o.Fill.BackgroundColor = Colors.Rgb(79, 129, 189)` |
| `cell.Style.Numberformat.Format = "#,##0.00"` | `o.NumberFormat = NumberFormats.Decimal2` |
| `ws.Column(c).Width = 15` | `ws.Column(c).SetWidth(15)` |
| `ws.View.FreezePanes(2, 1)` | `ws.FreezePanes(rows: 1)` |
| `ws.Cells[...].AutoFitColumns()` | Not available (needs font metrics); set widths explicitly |

## 🛠️ Architecture

SheetStream splits spreadsheet generation into low-overhead components:

1. **Package & Zip Archive Manager (`ExcelPackage`):** Owns the `ZipArchive`, streams each worksheet part as it is written, and on save emits `[Content_Types].xml`, the `.rels` files, `workbook.xml`, `styles.xml` and `sharedStrings.xml`.
2. **Sequential XmlWriter Engine (`Worksheet` / `RowWriter`):** Each row goes straight through an `XmlWriter` into the compressed stream. Column letters are cached and the row number is formatted once per row, so writing a cell allocates nothing.
3. **The Shared String Table (`SharedStringTable`):** Aggregates and de-duplicates string values across all sheets to minimise file size, exactly as Excel does.
4. **Styling Engine (`StyleSheet`):** De-duplicates fonts, fills, borders, number formats and cell formats into the five tables `styles.xml` requires, in the exact element order the schema demands.
5. **ECMA-376 Compliance:** Element ordering, mandatory placeholder fills, `_xHHHH_` escaping of control characters and the other details that decide whether Excel opens a file silently or shows a "repair" dialog are handled for you and covered by schema-validation tests.

## 🧪 Building and testing

```bash
dotnet build
dotnet test
dotnet run --project samples/SheetStream.Sample            # writes FinancialReport.xlsx
dotnet run -c Release --project samples/SheetStream.Sample -- 1000000   # 1M-row stress run
```

## 🤝 Contributing

We welcome the .NET community to help build the best free spreadsheet engine out there! See [CONTRIBUTING.md](CONTRIBUTING.md). Feel free to fork, open issues, submit pull requests, or optimize edge cases.

## 📄 License

This project is licensed under the **MIT License** - see the [LICENSE](LICENSE) file for details.
