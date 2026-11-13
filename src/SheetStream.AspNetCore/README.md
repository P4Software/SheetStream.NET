# SheetStream.AspNetCore

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) [![NuGet](https://img.shields.io/nuget/v/SheetStream.AspNetCore.svg)](https://www.nuget.org/packages/SheetStream.AspNetCore) [![Downloads](https://img.shields.io/nuget/dt/SheetStream.AspNetCore.svg)](https://www.nuget.org/packages/SheetStream.AspNetCore) [![Build](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml/badge.svg)](https://github.com/P4Software/SheetStream.NET/actions/workflows/build.yml) [![Docs](https://img.shields.io/badge/docs-p4software.github.io-0A5FAD)](https://p4software.github.io/SheetStream.NET/)

Return an Excel (`.xlsx`) download from an ASP.NET Core endpoint in one line. This package is the web companion of [SheetStream.NET](https://www.nuget.org/packages/SheetStream.NET), a zero-dependency, MIT-licensed library that writes and reads `.xlsx` files by streaming.

It exists because sending a generated workbook from a web endpoint has three traps: ASP.NET Core refuses synchronous writes to the response by default, building the file in a `MemoryStream` holds the whole export in memory, and a file name with accents or `ñ` needs special header encoding. `ExcelResults.Excel` handles all three.

## Install

```bash
dotnet add package SheetStream.AspNetCore
```

Targets `net8.0`. It brings in SheetStream.NET and nothing else; ASP.NET Core itself comes from the shared framework.

## Use it

```csharp
using SheetStream;
using SheetStream.AspNetCore;

// Minimal API, data already in memory
app.MapGet("/products.xlsx", (ProductService products) =>
    ExcelResults.Excel("products.xlsx", wb => wb.Worksheets.Add("Products").LoadFromCollection(products.All())));

// Minimal API, streamed from EF Core without loading the table into memory
app.MapGet("/orders.xlsx", (AppDb db) =>
    ExcelResults.Excel("orders.xlsx", async (wb, ct) =>
        await wb.Worksheets.Add("Orders").LoadFromAsyncEnumerable(
            db.Orders.AsNoTracking().AsAsyncEnumerable(),
            new LoadOptions { TableName = "Orders" }, ct)));

// MVC controller (.NET 7 and later execute IResult return values)
[HttpGet("export")]
public IResult Export() => ExcelResults.Excel("report.xlsx", wb => BuildReport(wb));

// Alongside the built-in results
app.MapGet("/summary.xlsx", () => Results.Extensions.Excel("summary.xlsx", wb => BuildSummary(wb)));
```

Inside the delegate you have the full SheetStream.NET workbook API: styles, tables, charts, conditional formatting, several sheets. You do not call `Save()`; the result finishes the file after your delegate returns.

```csharp
app.MapGet("/sales.xlsx", (SalesService sales) => ExcelResults.Excel("sales.xlsx", wb =>
{
    var money = wb.Styles.Create(o => o.NumberFormat = "#,##0.00");
    var sheet = wb.Worksheets.Add("Sales");
    sheet.AutoFitColumns();
    sheet.FreezePanes(rows: 1);
    sheet.Row().WriteCells("Region", "Month", "Revenue");
    foreach (var line in sales.ByRegionAndMonth())
        sheet.Row().WriteCell(line.Region).WriteCell(line.Month).WriteCell(line.Revenue, money);
}));
```

## What the response looks like

| | |
|---|---|
| Content type | `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` (also available as `ExcelResult.ContentType`) |
| File name | Sent as an attachment. `.xlsx` is added when the name you pass does not end in it. Non-ASCII names are sent through `filename*`, so `Reporte año 2026` arrives intact |
| `Content-Length` | Always set, because the file is complete before the first byte is sent. Browsers show real download progress |
| Memory | Flat. The workbook is written to a temporary file, then copied to the response asynchronously; the file is deleted when the request ends |
| Cancellation | The asynchronous overload passes you the request's cancellation token. Hand it to your query so a closed browser tab stops the export |
| Errors | An exception in your delegate happens before anything is sent, so the client gets your normal error response, not a truncated file |

## Options

Both overloads take an optional `PackageOptions`, the same class the core library uses. The setting most worth knowing for web exports:

```csharp
// About twice as fast to produce, files roughly 20% larger.
ExcelResults.Excel("big.xlsx", build, new PackageOptions { CompressionLevel = CompressionLevel.Fastest });
```

## When to use something else

If you need the bytes themselves (to attach to an e-mail or store in a database) use `ExcelPackage.CreateBytes` or `CreateBytesAsync` from the core package. On frameworks older than .NET 8, build the workbook with the core package into a stream and return it with your framework's own file result.

## More

Documentation and API reference: https://p4software.github.io/SheetStream.NET/. A runnable example with a download button is in [samples/SheetStream.Sample.Web](https://github.com/P4Software/SheetStream.NET/tree/main/samples/SheetStream.Sample.Web). Source, issues and the security policy: https://github.com/P4Software/SheetStream.NET.

MIT licensed, published by P4 Software.
