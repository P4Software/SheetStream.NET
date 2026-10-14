# SheetStream.AspNetCore

ASP.NET Core helpers for [SheetStream](https://www.nuget.org/packages/SheetStream), the MIT-licensed, zero-dependency streaming `.xlsx` library.

```bash
dotnet add package SheetStream.AspNetCore
```

## Return an Excel download in one line

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

The workbook is built into a temporary file and copied to the response asynchronously, so memory stays flat for large exports, the response has an accurate `Content-Length`, and no synchronous I/O touches the response body. Non-ASCII file names are sent correctly through `filename*`.

MIT licensed, published by P4 Software.
