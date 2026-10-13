using System;
using System.Diagnostics;
using System.IO;
using SheetStream;
using SheetStream.Styles;

// Usage example: a 3-column, 100-row styled data table saved to disk.
// Run with `dotnet run --project samples/SheetStream.Sample` and open FinancialReport.xlsx.

var path = Path.Combine(Environment.CurrentDirectory, "FinancialReport.xlsx");
var stopwatch = Stopwatch.StartNew();

using (var package = ExcelPackage.Create(path))
{
    var workbook = package.Workbook;
    var sheet = workbook.Worksheets.Add("Q1 Report");

    // Layout settings go first: the OpenXML schema places <cols> and frozen panes before the sheet data,
    // and SheetStream writes the sheet forward-only.
    sheet.Column(1).SetWidth(10);
    sheet.Column(2).SetWidth(28);
    sheet.Column(3).SetWidth(18);
    sheet.FreezePanes(rows: 1);
    sheet.AutoFilter(headerRow: 1, firstColumn: 1, lastColumn: 3);

    // Styles are registered once and de-duplicated; pass them to any cell.
    var header = workbook.Styles.Create(o =>
    {
        o.Font.Bold = true;
        o.Font.Color = Colors.White;
        o.Fill.BackgroundColor = Colors.ExcelBlue;
        o.Border.Style = BorderStyle.Thin;
        o.HorizontalAlignment = HorizontalAlignment.Center;
    });
    var currency = workbook.Styles.Create(o => o.NumberFormat = "$#,##0.00");
    var currencyStriped = workbook.Styles.Create(o =>
    {
        o.NumberFormat = "$#,##0.00";
        o.Fill.BackgroundColor = Colors.LightGray;
    });
    var striped = workbook.Styles.Create(o => o.Fill.BackgroundColor = Colors.LightGray);
    var total = workbook.Styles.Create(o =>
    {
        o.Font.Bold = true;
        o.NumberFormat = "$#,##0.00";
        o.Border.Top = BorderStyle.Double;
    });

    // Header row.
    sheet.Row(1)
        .WriteCell("ID", header)
        .WriteCell("Product Name", header)
        .WriteCell("Revenue (USD)", header);

    // 100 data rows, zebra-striped. Strings go through the shared string table automatically.
    for (var i = 1; i <= 100; i++)
    {
        var even = i % 2 == 0;
        sheet.Row(i + 1)
            .WriteCell(i, even ? striped : null)
            .WriteCell($"Enterprise SKU-{i:000}", even ? striped : null)
            .WriteCell(1500.50 * i, even ? currencyStriped : currency);
    }

    // Totals row with a live formula; Excel calculates it on open.
    sheet.Row(102)
        .WriteCell("Total", total)
        .WriteEmpty(total)
        .WriteFormula("SUM(C2:C101)", total);

    // A second sheet shows dates and mixed types written through the object overload.
    var meta = workbook.Worksheets.Add("Metadata");
    meta.Column(1).SetWidth(20);
    meta.Column(2).SetWidth(24);
    meta.Row().WriteCells("Generated", DateTime.Now);
    meta.Row().WriteCells("Rows", 100);
    meta.Row().WriteCells("Verified", true);
    meta.Row().WriteCells("Build time", TimeSpan.FromMinutes(1.5));

    package.Save();
}

stopwatch.Stop();
Console.WriteLine($"Wrote {path}");
Console.WriteLine($"Size: {new FileInfo(path).Length:N0} bytes in {stopwatch.ElapsedMilliseconds} ms");

// Optional stress run: `dotnet run -c Release --project samples/SheetStream.Sample -- 1000000`
if (args.Length > 0 && int.TryParse(args[0], out var rows) && rows > 0)
{
    var bigPath = Path.Combine(Environment.CurrentDirectory, "Stress.xlsx");
    var sw = Stopwatch.StartNew();
    using (var package = ExcelPackage.Create(bigPath, new PackageOptions { CompressionLevel = System.IO.Compression.CompressionLevel.Fastest }))
    {
        var sheet = package.Workbook.Worksheets.Add("Data");
        var money = package.Workbook.Styles.Create(o => o.NumberFormat = NumberFormats.Decimal2);
        sheet.Row(1).WriteCells("Id", "Category", "Amount", "When");
        var start = new DateTime(2026, 1, 1);
        for (var i = 1; i <= rows; i++)
        {
            sheet.Row()
                .WriteCell(i)
                .WriteCell(i % 7 == 0 ? "Hardware" : i % 3 == 0 ? "Services" : "Software")
                .WriteCell(i * 1.25, money)
                .WriteCell(start.AddMinutes(i));
        }
    }
    sw.Stop();
    Console.WriteLine($"Stress: {rows:N0} rows x 4 cols in {sw.Elapsed.TotalSeconds:F1}s, " +
                      $"{new FileInfo(bigPath).Length / 1024 / 1024:N0} MB, peak working set {Process.GetCurrentProcess().PeakWorkingSet64 / 1024 / 1024:N0} MB");
}
