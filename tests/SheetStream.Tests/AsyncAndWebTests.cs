using System.Linq;
using System.Runtime.CompilerServices;
using Microsoft.AspNetCore.Http;
using SheetStream.AspNetCore;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>Async loading and the ASP.NET Core download result.</summary>
public class AsyncAndWebTests
{
    public sealed class Item
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
    }

    private static async IAsyncEnumerable<Item> Items(int count, [EnumeratorCancellation] CancellationToken ct = default)
    {
        for (var i = 1; i <= count; i++)
        {
            ct.ThrowIfCancellationRequested();
            if (i % 100 == 0) await Task.Yield(); // Behave like a real async source.
            yield return new Item { Id = i, Name = "Item " + i };
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(255)]
    [InlineData(256)]
    [InlineData(1000)]
    public async Task LoadFromAsyncEnumerable_WritesEveryItemOnce(int count)
    {
        LoadResult result = default;
        var bytes = await Task.Run(async () =>
        {
            using var ms = new MemoryStream();
            using (var p = new ExcelPackage(ms, leaveOpen: true))
            {
                result = await p.Workbook.Worksheets.Add("A").LoadFromAsyncEnumerable(Items(count), new LoadOptions { TableName = "Items" });
                p.Save();
            }
            return ms.ToArray();
        });

        Assert.Equal(count, result.DataRows);
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("A").ToList();
        Assert.Equal(new object?[] { "Id", "Name" }, rows[0].Values);
        Assert.Equal(Enumerable.Range(1, count).Select(i => (object?)(double)i), rows.Skip(1).Where(r => r.Values.Count > 0).Select(r => r[0]));
        using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2016).Validate(doc));
    }

    [Fact]
    public async Task LoadFromAsyncEnumerable_HonoursCancellation()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        using var p = new ExcelPackage(new MemoryStream());
        var sheet = p.Workbook.Worksheets.Add("A");
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => sheet.LoadFromAsyncEnumerable(Items(10), null, cts.Token));
    }

    [Fact]
    public async Task ExcelResult_StreamsAValidDownload()
    {
        var context = new DefaultHttpContext();
        var body = new MemoryStream();
        context.Response.Body = body;

        var result = ExcelResults.Excel("Reporte año 2026", async (wb, ct) =>
            await wb.Worksheets.Add("Items").LoadFromAsyncEnumerable(Items(50), null, ct));
        await result.ExecuteAsync(context);

        Assert.Equal(ExcelResult.ContentType, context.Response.ContentType);
        Assert.Equal(body.Length, context.Response.ContentLength);
        var disposition = context.Response.Headers["Content-Disposition"].ToString();
        Assert.Contains("attachment", disposition);
        Assert.Contains("filename*=UTF-8''Reporte%20a%C3%B1o%202026.xlsx", disposition);

        using var reader = ExcelReader.Open(new MemoryStream(body.ToArray()));
        Assert.Equal(51, reader.ReadRows("Items").Count());
    }

    [Fact]
    public async Task ResultsExtensions_BuildsSynchronously()
    {
        var context = new DefaultHttpContext { Response = { Body = new MemoryStream() } };
        await Results.Extensions.Excel("x", wb => wb.Worksheets.Add("S").Row().WriteCell(1)).ExecuteAsync(context);
        Assert.True(context.Response.ContentLength > 0);
    }
}
