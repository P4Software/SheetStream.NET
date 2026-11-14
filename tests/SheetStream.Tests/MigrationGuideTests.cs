using System.Linq;
using SheetStream.Reading;
using SheetStream.Styles;

namespace SheetStream.Tests;

/// <summary>Keeps the code in docs/migrating-from-epplus.md honest: it must compile and produce the right file.</summary>
public class MigrationGuideTests
{
    public sealed class QueryResultDto
    {
        public List<string> Columns { get; set; } = new();
        public List<Dictionary<string, object?>> Rows { get; set; } = new();
    }

    // Copied verbatim from the guide's "After" section.
    public Task<byte[]> ExportToExcelAsync(QueryResultDto result)
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var header = wb.Styles.Create(o =>
            {
                o.Font.Bold = true;
                o.Font.Color = Colors.White;
                o.Fill.BackgroundColor = Colors.Rgb(79, 129, 189);
                o.HorizontalAlignment = HorizontalAlignment.Center;
            });
            var stripe = wb.Styles.Create(o => o.Fill.BackgroundColor = Colors.Rgb(242, 242, 242));
            Func<string, bool, CellStyle> fmt = (code, striped) =>
                wb.Styles.Create(o => { o.NumberFormat = code; if (striped) o.Fill.BackgroundColor = Colors.Rgb(242, 242, 242); });
            var dateTime = new[] { fmt("yyyy-mm-dd hh:mm:ss", false), fmt("yyyy-mm-dd hh:mm:ss", true) };
            var money = new[] { fmt("#,##0.00", false), fmt("#,##0.00", true) };
            var whole = new[] { fmt("#,##0", false), fmt("#,##0", true) };

            var sheet = wb.Worksheets.Add("Query Results");
            sheet.AutoFitColumns();
            sheet.FreezePanes(rows: 1);

            var head = sheet.Row(1);
            foreach (var column in result.Columns) head.WriteCell(column, header);

            for (var r = 0; r < result.Rows.Count; r++)
            {
                var striped = r % 2 == 1 ? 1 : 0;
                var row = sheet.Row();
                foreach (var column in result.Columns)
                {
                    result.Rows[r].TryGetValue(column, out var value);
                    switch (value)
                    {
                        case DateTime d: row.WriteCell(d, dateTime[striped]); break;
                        case decimal or double or float: row.WriteCell(value, money[striped]); break;
                        case int or long: row.WriteCell(value, whole[striped]); break;
                        case null: row.WriteEmpty(striped == 1 ? stripe : null); break;
                        default: row.WriteCell(value.ToString(), striped == 1 ? stripe : null); break;
                    }
                }
            }
        });
        return Task.FromResult(bytes);
    }

    [Fact]
    public async Task GuideExample_ProducesTheSameDataAsTheEpplusVersion()
    {
        var dto = new QueryResultDto { Columns = { "Customer", "Total", "Orders", "Last order" } };
        for (var i = 0; i < 5; i++)
            dto.Rows.Add(new() { ["Customer"] = "C" + i, ["Total"] = 10.5m * i, ["Orders"] = i, ["Last order"] = new DateTime(2026, 1, 1 + i) });
        dto.Rows.Add(new() { ["Customer"] = "Missing values" });

        var bytes = await ExportToExcelAsync(dto);
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("Query Results").ToList();
        Assert.Equal(new object?[] { "Customer", "Total", "Orders", "Last order" }, rows[0].Values);
        Assert.Equal(42.0, rows[5][1]);
        Assert.Equal(new DateTime(2026, 1, 5), rows[5][3]);
        Assert.Equal("Missing values", rows[6][0]);
    }

    // ---------------------------------------------------------------- docs/migrating-from-closedxml.md

    public sealed class Order
    {
        public int Id { get; set; }
        public string Customer { get; set; } = "";
        public int Quantity { get; set; }
        public double Amount { get; set; }
        public DateTime Date { get; set; }
    }

    // Copied verbatim from the ClosedXML guide's "After" section. The "Before" code there is the ClosedXML
    // writer of the benchmark project in shortened form, which is compiled against ClosedXML itself.
    public void WriteOrders(Stream output, Order[] orders)
    {
        using var package = new ExcelPackage(output, leaveOpen: true);
        var workbook = package.Workbook;

        // Styles are made once and passed with the cell. Nothing is restyled afterwards.
        var bold = workbook.Styles.Create(o => o.Font.Bold = true);
        var money = workbook.Styles.Create(o => o.NumberFormat = "#,##0.00");
        var date = workbook.Styles.Create(o => o.NumberFormat = "yyyy-mm-dd");

        var sheet = workbook.Worksheets.Add("Orders");
        sheet.AutoFitColumns();          // Replaces Columns().AdjustToContents() at the end
        sheet.FreezePanes(rows: 1);      // Replaces SheetView.FreezeRows(1)

        sheet.Row(1).WriteCells(bold, new[] { "Id", "Customer", "Quantity", "Amount", "Date" });
        foreach (var o in orders)
            sheet.Row().WriteCell(o.Id).WriteCell(o.Customer).WriteCell(o.Quantity).WriteCell(o.Amount, money).WriteCell(o.Date, date);

        package.Save();                  // Nothing is written until Save
    }

    [Fact]
    public void ClosedXmlGuide_AfterCode_WritesTheSheet_AndItsOtherSnippetsCompile()
    {
        var orders = Enumerable.Range(1, 20).Select(i => new Order { Id = i, Customer = "Customer " + i, Quantity = i * 2, Amount = i * 10.25, Date = new DateTime(2026, 3, i) }).ToArray();
        var output = new MemoryStream();
        WriteOrders(output, orders);

        // The reading snippet from the guide.
        using var reader = ExcelReader.Open(new MemoryStream(output.ToArray()));
        var seen = 0;
        foreach (var row in reader.ReadRows("Orders"))
        {
            if (row.Index == 1) { Assert.Equal("Customer", row.GetString(1)); continue; }
            var id = row.GetDouble(0);            // Columns are zero-based here
            var customer = row.GetString(1);
            Assert.Equal("Customer " + (int)id!, customer);
            Assert.Equal(new DateTime(2026, 3, (int)id!), row.GetDateTime(4));
            seen++;
        }
        Assert.Equal(20, seen);

        // The one-call form, and every right-hand cell of the quick-reference table.
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var sheet = wb.Worksheets.Add("Orders");
            sheet.LoadFromCollection(orders, new LoadOptions { TableName = "Orders" });

            var styled = wb.Worksheets.Add("Reference");
            var style = wb.Styles.Create(o =>
            {
                o.Font.Bold = true;
                o.Fill.BackgroundColor = "#00FFFF";
                o.Font.Color = "#00008B";
                o.HorizontalAlignment = HorizontalAlignment.Center;
                o.NumberFormat = "#,##0.00";
            });
            styled.AutoFitColumns();
            styled.Column(2).SetWidth(20);
            styled.FreezePanes(rows: 1);
            styled.Row(1).WriteCell(1.5, style).WriteFormula("SUM(A1:A9)").WriteHyperlink("site", "https://example.com/");
            styled.MergeCells("A3:D3");
            styled.AutoFilter(1, 1, 4);
            styled.AddDataValidation("A5").List("Yes", "No");
            styled.Protect("password");

            wb.Worksheets.Add("Random").Cells[3, 2].Value = "any order";
        });
        using var check = ExcelReader.Open(new MemoryStream(bytes));
        Assert.Equal(21, check.ReadRows("Orders").Count());

        // The header call in the guide passes the style and then a collection of texts. A style followed by
        // loose arguments binds to the values-only overload; it must fail loudly, never write the style into
        // the sheet as text. And a literal null among the values stays what it always was, an empty cell.
        var styledHeaders = ExcelPackage.CreateBytes(wb =>
        {
            var bold = wb.Styles.Create(o => o.Font.Bold = true);
            var sheet = wb.Worksheets.Add("H");
            sheet.Row().WriteCells(bold, new[] { "a", "b" });
            sheet.Row().WriteCells(bold, new List<string> { "a", "b" });
            sheet.Row().WriteCells(bold, new object?[] { "a", "b" });
            sheet.Row().WriteCells("a", "b");
            sheet.Row().WriteCells(null, "a", "b");
            var ex = Assert.Throws<ArgumentException>(() => sheet.Row().WriteCells(bold, "a", "b"));
            Assert.Contains("WriteCells(style, new[]", ex.Message);
            Assert.Throws<ArgumentException>(() => sheet.Row().WriteCell((object)bold));
        });
        using var headers = ExcelReader.Open(new MemoryStream(styledHeaders));
        var written = headers.ReadRows("H").ToList();
        Assert.All(written.Take(4), row => Assert.Equal(new object?[] { "a", "b" }, row.Values));
        Assert.Equal(new object?[] { null, "a", "b" }, written[4].Values);
        Assert.Equal("any order", check.ReadRows("Random").Single()[1]);
    }
}
