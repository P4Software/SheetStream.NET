# SheetStream.AspNetCore

ASP.NET Core helpers for [SheetStream.NET](https://www.nuget.org/packages/SheetStream.NET), a zero-dependency, MIT-licensed library that writes and reads `.xlsx` files. It targets `net8.0`.

```bash
dotnet add package SheetStream.AspNetCore
```

## Return an Excel download

```csharp
using SheetStream;
using SheetStream.AspNetCore;

// Minimal API, synchronous data
app.MapGet("/products.xlsx", (ProductService products) =>
    ExcelResults.Excel("products.xlsx", wb => wb.Worksheets.Add("Products").LoadFromCollection(products.All())));

// Minimal API, streamed from EF Core without loading the table into memory
app.MapGet("/orders.xlsx", (AppDb db) =>
    ExcelResults.Excel("orders.xlsx", async (wb, ct) =>
        await wb.Worksheets.Add("Orders").LoadFromAsyncEnumerable(
            db.Orders.AsNoTracking().AsAsyncEnumerable(),
            new LoadOptions { TableName = "Orders" }, ct)));

// MVC controller (.NET 7+ executes IResult return values)
[HttpGet("export")]
public IResult Export() => ExcelResults.Excel("report.xlsx", wb => BuildReport(wb));
```

The workbook is built into a temporary file and copied to the response asynchronously. Memory stays flat for large exports, the response carries an accurate `Content-Length`, and nothing writes to the response body synchronously. Non-ASCII file names are sent through `filename*`.

You do not call `Save()` here: `ExcelResults.Excel` finishes the file after your build delegate returns.

Documentation: https://p4software.github.io/SheetStream.NET/. MIT licensed, published by P4 Software.
