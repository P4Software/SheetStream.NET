# SheetStream web sample

A small ASP.NET Core minimal API that returns Excel downloads with [SheetStream](../../README.md) and
[SheetStream.AspNetCore](../../src/SheetStream.AspNetCore/README.md). It exists to be copied from.

## What it shows

| Endpoint | Pattern |
|---|---|
| `GET /orders.xlsx?rows=N` | A large export (50,000 rows by default, any N up to Excel's limit) streamed from a simulated async database through `LoadFromAsyncEnumerable` into a styled table: bold header, number and date formats via `[SheetColumn]`, filter buttons, frozen header row, column widths, and a live `SUM` formula under the table. Memory stays flat whatever N is. |
| `GET /report.xlsx` | A small summary written with random-access `Cells[...]` (title, header row, totals with formulas) plus a column chart and a line chart that read those cells. |
| `GET /` | `wwwroot/index.html`: a page with the two download buttons and a row-count input. |

Everything is in [`Program.cs`](Program.cs); the lines worth copying are marked `COPY THIS`.

## How to run

Double-click `run.cmd`. It builds the project, starts the server at `http://localhost:5178` and opens the
browser. If something goes wrong, the window stays open with the message.

From a terminal:

```bash
dotnet run --project samples/SheetStream.Sample.Web --urls http://localhost:5178
```

Then open <http://localhost:5178>, or fetch a file directly:

```bash
curl -sS -o orders.xlsx "http://localhost:5178/orders.xlsx?rows=100000"
```

Requires the .NET 8 SDK. No NuGet packages beyond the two SheetStream projects it references.

## The six lines to copy into your own app

```csharp
using SheetStream;
using SheetStream.AspNetCore;

app.MapGet("/orders.xlsx", (AppDb db) =>
    ExcelResults.Excel("orders.xlsx", async (wb, ct) =>
        await wb.Worksheets.Add("Orders").LoadFromAsyncEnumerable(
            db.Orders.AsNoTracking().AsAsyncEnumerable(), new LoadOptions { TableName = "Orders" }, ct)));
```

Install with `dotnet add package SheetStream.AspNetCore` (it brings `SheetStream` with it). `ExcelResults.Excel`
builds the workbook into a temporary file and copies it to the response asynchronously, so the response has an
exact `Content-Length`, no synchronous I/O touches the body, and a 1,000,000-row export costs the same memory as
a 1,000-row one. The same result type is returned from an MVC action on .NET 7+ (`public IResult Export() => ...`).

Two rules the sample follows that are easy to miss:

- Column widths, frozen panes and other layout settings must be set **before the first row is written**;
  SheetStream writes forward-only.
- A sheet uses either `Row()`/`LoadFrom...` streaming or `Cells[...]` random access, not both. Stream the big
  data, use `Cells` for the small shaped summary.

## Calling it from Blazor

No Blazor project is needed to use these endpoints from Blazor. Keep the endpoint exactly as above in the host
`Program.cs` (Blazor Server, Blazor Web App or the server project of a hosted WebAssembly app) and make the
button navigate to it with a full page load, so the browser treats the response as a file download instead of
a Blazor route:

```razor
@inject NavigationManager Nav

<button class="btn btn-primary" @onclick="Download">Download orders.xlsx</button>

@code {
    private void Download() => Nav.NavigateTo("/orders.xlsx", forceLoad: true);
}
```

`forceLoad: true` is the whole trick: because the response carries `Content-Disposition: attachment`, the
browser saves the file and stays on the current page. Pass parameters in the query string
(`Nav.NavigateTo($"/orders.xlsx?rows={rows}", forceLoad: true)`); the current authentication cookie travels with
the request, so `[Authorize]` on the endpoint works as usual. A plain `<a href="/orders.xlsx" download>` inside
a Razor component does the same thing without any code.
