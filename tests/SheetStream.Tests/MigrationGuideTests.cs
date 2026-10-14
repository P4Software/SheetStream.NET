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
}
