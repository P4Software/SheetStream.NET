using System.IO;
using System.Threading.Tasks;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// LoadFromArrays (the EPPlus Cells.LoadFromArrays replacement) and CreateBytesAsync, both requested by the
/// first production migration.
/// </summary>
public class LoadFromArraysTests
{
    private static readonly object?[][] Rows =
    {
        new object?[] { "Sku", "Description", "Qty" },
        new object?[] { "A-1", "Bolt", 3 },
        new object?[] { "B-2", "Nut" },                 // Shorter: Qty stays empty.
        new object?[] { "C-3", "Washer", 7, "extra" },  // Longer: the extra value is dropped.
    };

    [Fact]
    public void FirstArrayIsTheHeaderAndTheRestAreRows()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var bold = wb.Styles.Create(o => o.Font.Bold = true);
            var result = wb.Worksheets.Add("Stock").LoadFromArrays(Rows, new LoadOptions { HeaderStyle = bold, TableName = "Stock" });
            Assert.Equal(1, result.FirstRow);
            Assert.Equal(4, result.LastRow);
            Assert.Equal(3, result.LastColumn);
            Assert.Equal(3, result.DataRows);
            Assert.Equal(new[] { "Sku", "Description", "Qty" }, result.Headers);
            Assert.NotNull(result.Table);
        });

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("Stock").ToList();
        Assert.Equal(4, rows.Count);
        Assert.Equal("Sku", rows[0].GetString(0));
        Assert.Equal(3.0, rows[1].GetDouble(2));
        Assert.Null(rows[2][2]);
        Assert.Equal(3, rows[3].Values.Count); // "extra" was not written.
    }

    [Fact]
    public void WithoutHeadersTheFirstArrayIsData()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
            wb.Worksheets.Add("Data").LoadFromArrays(Rows, new LoadOptions { PrintHeaders = false, StartColumn = 2 }));

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("Data").ToList();
        Assert.Equal(4, rows.Count);
        Assert.Null(rows[0][0]);                      // Column A is skipped (StartColumn = 2).
        Assert.Equal("Sku", rows[0].GetString(1));    // The "header" array is an ordinary row here.
        Assert.Equal("Bolt", rows[1].GetString(2));
    }

    [Fact]
    public void EmptyInputWritesNothing()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var result = wb.Worksheets.Add("Empty").LoadFromArrays(System.Array.Empty<object?[]>());
            Assert.Equal(0, result.DataRows);
        });
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        Assert.Empty(reader.ReadRows("Empty"));
    }

    [Fact]
    public async Task CreateBytesAsyncAwaitsBetweenSheets()
    {
        var bytes = await ExcelPackage.CreateBytesAsync(async wb =>
        {
            wb.Worksheets.Add("First").LoadFromArrays(await FakeQueryAsync("one"));
            wb.Worksheets.Add("Second").LoadFromArrays(await FakeQueryAsync("two"));
        });

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        Assert.Equal(new[] { "First", "Second" }, reader.SheetNames);
        Assert.Equal("two", reader.ReadRows("Second").Skip(1).First().GetString(0));
    }

    private static async Task<object?[][]> FakeQueryAsync(string value)
    {
        await Task.Yield(); // Really asynchronous, so the builder resumes on another continuation.
        return new[] { new object?[] { "Value" }, new object?[] { value } };
    }
}
