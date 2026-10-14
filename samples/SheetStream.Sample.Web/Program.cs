// SheetStream web sample: two Excel download endpoints in a minimal API.
//
//   GET /orders.xlsx?rows=N  streams N generated orders (default 50,000) from a fake async database
//                            into a styled table, without ever holding the whole data set in memory.
//   GET /report.xlsx         builds a small summary with random-access Cells[...] and places a chart.
//
// Run with `dotnet run --urls http://localhost:5178` (or double-click run.cmd) and open the page.
// Everything a reader needs to copy into their own app is marked "COPY THIS".

using System.Globalization;
using System.Runtime.CompilerServices;
using SheetStream;                 // COPY THIS: the workbook API (Worksheets, Styles, LoadFromAsyncEnumerable, charts)
using SheetStream.AspNetCore;      // COPY THIS: ExcelResults.Excel(...) returns the download
using SheetStream.Styles;          // Colors and NumberFormats constants

var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

// The landing page is a static file: wwwroot/index.html, served at "/".
app.UseDefaultFiles();
app.UseStaticFiles();

// Excel's hard limit is 1,048,576 rows per sheet; one row is taken by the header.
const int MaxRows = 1_048_575;
const int DefaultRows = 50_000;

// ------------------------------------------------------------------------------------------------------
// Pattern 1: a large export streamed from an async source (EF Core, Dapper, a paged API, ...).
// ------------------------------------------------------------------------------------------------------
app.MapGet("/orders.xlsx", (int? rows) =>
{
    // Minimal APIs bind `?rows=N` to the int? parameter; anything out of range is clamped rather than rejected
    // so a typo never produces an error page where a person expected a download.
    var count = Math.Clamp(rows ?? DefaultRows, 1, MaxRows);

    // COPY THIS: ExcelResults.Excel builds the workbook into a temp file and copies it to the response
    // asynchronously, so memory stays flat, Content-Length is exact and no synchronous I/O touches the body.
    return ExcelResults.Excel("orders.xlsx", async (workbook, ct) =>
    {
        var sheet = workbook.Worksheets.Add("Orders");

        // Layout settings go FIRST. SheetStream writes the sheet forward-only and the OpenXML schema places
        // column widths and frozen panes before the row data, so these calls throw once a row has been written.
        sheet.Column(1).SetWidth(10);   // Order #
        sheet.Column(2).SetWidth(12);   // Date
        sheet.Column(3).SetWidth(28);   // Customer
        sheet.Column(4).SetWidth(24);   // Product
        sheet.Column(5).SetWidth(8);    // Qty
        sheet.Column(6).SetWidth(12);   // Unit price
        sheet.Column(7).SetWidth(14);   // Total
        sheet.Column(8).SetWidth(12);   // Status
        sheet.FreezePanes(rows: 1);     // Keep the header visible while scrolling.

        // Styles are registered once per workbook and de-duplicated; pass the handle to any cell.
        var header = workbook.Styles.Create(o =>
        {
            o.Font.Bold = true;
            o.Font.Color = Colors.White;
            o.Fill.BackgroundColor = "#0A5FAD"; // P4 Software blue
        });

        var options = new LoadOptions
        {
            HeaderStyle = header,
            // TableName turns the range into an Excel table: banded rows and filter buttons on every header.
            TableName = "Orders",
            TableStyle = "TableStyleMedium2",
        };
        // Number and date formats come from the [SheetColumn] attributes on the Order class below. The
        // alternative, when you cannot annotate the model, is options.ColumnStyles["Total"] = someStyle.

        // COPY THIS: LoadFromAsyncEnumerable awaits the source between rows and writes each row straight
        // into the zip entry. Swap FakeDatabase.ReadOrdersAsync for db.Orders.AsNoTracking().AsAsyncEnumerable().
        var result = await sheet.LoadFromAsyncEnumerable(FakeDatabase.ReadOrdersAsync(count, ct), options, ct);

        // A totals row under the table. Excel evaluates the formula when the file opens, so the server does no
        // arithmetic and the number stays live when the user filters or edits.
        var totalStyle = workbook.Styles.Create(o =>
        {
            o.Font.Bold = true;
            o.NumberFormat = "$#,##0.00";
            o.Border.Top = BorderStyle.Double;
        });
        var label = workbook.Styles.Create(o => o.Font.Bold = true);
        sheet.Row(result.LastRow + 2)
            .WriteCell($"{result.DataRows:N0} orders", label)
            .Skip(5)
            .WriteFormula($"SUM(G{result.FirstRow + 1}:G{result.LastRow})", totalStyle);
    });
});

// ------------------------------------------------------------------------------------------------------
// Pattern 2: a small, shaped report written cell by cell, plus a chart that reads those cells.
// ------------------------------------------------------------------------------------------------------
app.MapGet("/report.xlsx", () =>
    // The synchronous overload is enough here: the data is already in memory.
    ExcelResults.Excel("report.xlsx", workbook =>
    {
        var sheet = workbook.Worksheets.Add("Summary");

        // Random-access mode: Cells[...] can be set in any order and are written sorted when the workbook is
        // saved. Memory grows with the cell count, so use it for summaries, not for the 50,000-row export above.
        sheet.Column(1).SetWidth(12);
        sheet.Column(2).SetWidth(16);
        sheet.Column(3).SetWidth(10);

        var title = workbook.Styles.Create(o => { o.Font.Bold = true; o.Font.Size = 14; o.Font.Color = "#0A5FAD"; });
        var header = workbook.Styles.Create(o =>
        {
            o.Font.Bold = true;
            o.Font.Color = Colors.White;
            o.Fill.BackgroundColor = "#0A5FAD";
            o.HorizontalAlignment = HorizontalAlignment.Center;
        });
        var money = workbook.Styles.Create(o => o.NumberFormat = "$#,##0");
        var count = workbook.Styles.Create(o => o.NumberFormat = NumberFormats.Integer);
        var totalMoney = workbook.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "$#,##0"; o.Border.Top = BorderStyle.Double; });
        var totalCount = workbook.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = NumberFormats.Integer; o.Border.Top = BorderStyle.Double; });

        sheet.Cells["A1"].Value = "Revenue by month, 2026";
        sheet.Cells["A1"].Style = title;

        sheet.Cells[3, 1].Value = "Month";   sheet.Cells[3, 1].Style = header;
        sheet.Cells[3, 2].Value = "Revenue"; sheet.Cells[3, 2].Style = header;
        sheet.Cells[3, 3].Value = "Orders";  sheet.Cells[3, 3].Style = header;

        // Aggregate the same generated orders the export uses, so the two downloads agree with each other.
        var byMonth = FakeDatabase.Orders(DefaultRows)
            .GroupBy(o => o.OrderDate.Month)
            .OrderBy(g => g.Key)
            .Select(g => (Month: g.Key, Revenue: g.Sum(o => o.Total), Count: g.Count()));

        var row = 4;
        foreach (var (month, revenue, orders) in byMonth)
        {
            // Month names as text (not dates) so the chart's category axis shows "Jan", "Feb", ... verbatim.
            sheet.Cells[row, 1].Value = CultureInfo.InvariantCulture.DateTimeFormat.GetAbbreviatedMonthName(month);
            sheet.Cells[row, 2].Value = revenue; sheet.Cells[row, 2].Style = money;
            sheet.Cells[row, 3].Value = orders;  sheet.Cells[row, 3].Style = count;
            row++;
        }
        var lastDataRow = row - 1;

        // Formulas work in random-access mode too; the leading "=" is optional.
        sheet.Cells[row, 1].Value = "Total";
        sheet.Cells[row, 2].Formula = $"SUM(B4:B{lastDataRow})"; sheet.Cells[row, 2].Style = totalMoney;
        sheet.Cells[row, 3].Formula = $"SUM(C4:C{lastDataRow})"; sheet.Cells[row, 3].Style = totalCount;

        // The chart fills the cell range given as placement and reads its data from the cells above when
        // Excel opens the file. A series name starting with "=" points at a cell; plain text is used literally.
        var chart = sheet.AddChart(ChartType.Column, "E3:M20");
        chart.Title = "Revenue by month";
        chart.Legend = LegendPosition.None;
        chart.YAxisTitle = "USD";
        chart.ValueAxisNumberFormat = "$#,##0";
        chart.AddSeries(values: $"B4:B{lastDataRow}", categories: $"A4:A{lastDataRow}", name: "=B3", color: "#0A5FAD");

        // A second chart on the same data, in P4 orange, to show that charts are cheap to add.
        var line = sheet.AddChart(ChartType.LineMarkers, "E22:M38");
        line.Title = "Orders per month";
        line.Legend = LegendPosition.None;
        line.AddSeries(values: $"C4:C{lastDataRow}", categories: $"A4:A{lastDataRow}", name: "=C3", color: "#FBAD48");
    }));

app.Run();

// ------------------------------------------------------------------------------------------------------
// The "database". In a real app this is your EF Core DbSet, Dapper query or paged HTTP client.
// ------------------------------------------------------------------------------------------------------

/// <summary>
/// One exported row. Public properties become columns in declaration order unless <c>Order</c> says
/// otherwise; <see cref="SheetColumnAttribute"/> supplies the header text and the Excel number format.
/// </summary>
public sealed class Order
{
    [SheetColumn("Order #")]
    public int Id { get; set; }

    [SheetColumn("Date", NumberFormat = NumberFormats.Date)]
    public DateTime OrderDate { get; set; }

    [SheetColumn("Customer")]
    public string Customer { get; set; } = "";

    [SheetColumn("Product")]
    public string Product { get; set; } = "";

    [SheetColumn("Qty", NumberFormat = NumberFormats.Integer)]
    public int Quantity { get; set; }

    [SheetColumn("Unit price", NumberFormat = "$#,##0.00")]
    public decimal UnitPrice { get; set; }

    [SheetColumn("Total", NumberFormat = "$#,##0.00")]
    public decimal Total { get; set; }

    [SheetColumn("Status")]
    public string Status { get; set; } = "";
}

/// <summary>
/// Generates deterministic sample orders. The async method stands in for a database driver: it yields rows
/// in batches and gives up the thread between batches, exactly where a real driver would await the network.
/// </summary>
public static class FakeDatabase
{
    private static readonly string[] Customers =
    {
        "Acme Logistics", "Barrdega Foods", "Contoso Retail", "Globex Pharma", "Initech Systems",
        "Northwind Traders", "Soylent Corp", "Umbrella Freight", "Vandelay Imports", "Wayne Industries",
    };

    private static readonly (string Name, decimal Price)[] Products =
    {
        ("Pallet jack", 349.00m), ("Barcode scanner", 189.50m), ("Label printer", 425.00m),
        ("Shrink wrap (case)", 64.99m), ("Steel shelving unit", 279.00m), ("Safety vest", 12.75m),
        ("Forklift battery", 1_890.00m), ("Hand truck", 98.00m),
    };

    private static readonly string[] Statuses = { "Shipped", "Shipped", "Shipped", "Pending", "Cancelled" };

    /// <summary>Synchronous generator used by the report; a fixed seed keeps every download identical.</summary>
    public static IEnumerable<Order> Orders(int count)
    {
        var random = new Random(2026);
        var start = new DateTime(2026, 1, 1);
        for (var i = 1; i <= count; i++)
        {
            var (name, price) = Products[random.Next(Products.Length)];
            var quantity = random.Next(1, 25);
            yield return new Order
            {
                Id = 10_000 + i,
                OrderDate = start.AddDays(random.Next(365)),
                Customer = Customers[random.Next(Customers.Length)],
                Product = name,
                Quantity = quantity,
                UnitPrice = price,
                Total = price * quantity,
                Status = Statuses[random.Next(Statuses.Length)],
            };
        }
    }

    /// <summary>
    /// The async source the export streams from. <c>[EnumeratorCancellation]</c> lets the cancellation token
    /// SheetStream passes to <c>GetAsyncEnumerator</c> reach this method, so closing the browser tab stops
    /// the generator as well as the download.
    /// </summary>
    public static async IAsyncEnumerable<Order> ReadOrdersAsync(int count, [EnumeratorCancellation] CancellationToken ct = default)
    {
        const int BatchSize = 500;
        var produced = 0;
        foreach (var order in Orders(count))
        {
            // Simulate one round trip to the database per batch. Task.Yield keeps the sample fast; a real
            // driver would await its network read here and the sheet writer would simply wait for it.
            if (produced % BatchSize == 0)
            {
                ct.ThrowIfCancellationRequested();
                await Task.Yield();
            }
            yield return order;
            produced++;
        }
    }
}
