using System.Data;
using System.Runtime.CompilerServices;
using SheetStream.Reading;
using SheetStream.Styles;

namespace SheetStream.Tests;

/// <summary>
/// The asynchronous reader (OpenAsync, ReadRowsAsync, ReadAsync&lt;T&gt;, ReadDataTableAsync). Every workbook is
/// written with SheetStream and the async results are compared with the synchronous API on the same bytes, so
/// the two parsers are proven to agree cell for cell.
/// </summary>
public class AsyncReaderTests
{
    public sealed class Product
    {
        public int Id { get; set; }
        public string Name { get; set; } = "";
        public double Price { get; set; }
        public bool InStock { get; set; }
        public DateTime? Shipped { get; set; }
    }

    /// <summary>A workbook with every value kind the reader types, gaps, blank rows and a second sheet.</summary>
    private static byte[] Workbook(int dataRows = 25) => ExcelPackage.CreateBytes(wb =>
    {
        var s = wb.Worksheets.Add("Products");
        s.Row().WriteCells("Id", "Name", "Price", "InStock", "Shipped");
        for (var i = 1; i <= dataRows; i++)
        {
            var r = s.Row().WriteCell(i).WriteCell("Product " + i + (i % 7 == 0 ? " & co <x>" : "")).WriteCell(i * 1.25).WriteCell(i % 2 == 0);
            if (i % 3 == 0) r.WriteEmpty(); else r.WriteCell(new DateTime(2026, 1, 1).AddDays(i));
            if (i % 5 == 0) s.Row(); // A blank row: ReadRows yields it empty, Read<T> skips it.
        }
        s.Row().Skip(5).WriteCell(TimeSpan.FromHours(1.5)).WriteCell("far right"); // Sparse row (F, G) outside the header: Read<T> skips it.
        wb.Worksheets.Add("Second").Row().WriteCells("a", 1.0);
    });

    /// <summary>Hides seekability, the way an HTTP request body or a network stream does.</summary>
    private sealed class ForwardOnlyStream(Stream inner) : Stream
    {
        public bool Disposed { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => inner.Read(buffer, offset, count);
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken ct = default)
        {
            await Task.Yield(); // Behave like a real network stream: complete asynchronously.
            return await inner.ReadAsync(buffer.Slice(0, Math.Min(buffer.Length, 1000)), ct); // Small chunks exercise the loop.
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { Disposed = true; base.Dispose(disposing); }
    }

    /// <summary>A hand-written mapper, as a user without the source generator would write one.</summary>
    private sealed class ProductMapper : ISheetRowMapper<Product>
    {
        public IReadOnlyList<string> Headers { get; } = new[] { "Id", "Name", "Price" };
        public IReadOnlyList<string?> NumberFormats { get; } = new string?[3];
        public bool CanRead => true;
        public Type MemberType(int column) => column switch { 0 => typeof(int), 1 => typeof(string), _ => typeof(double) };
        public bool CanSet(int column) => true;
        public int FindColumn(string header)
        {
            for (var i = 0; i < Headers.Count; i++)
                if (string.Equals(Headers[i], header, StringComparison.OrdinalIgnoreCase)) return i;
            return -1;
        }
        public void WriteRow(Product item, RowWriter row, CellStyle?[] styles) => row.WriteCell(item.Id).WriteCell(item.Name).WriteCell(item.Price);
        public object? GetValue(Product item, int column) => column switch { 0 => item.Id, 1 => item.Name, _ => item.Price };
        public Product Create() => new();
        public void SetValue(ref Product item, int column, object? value)
        {
            switch (column)
            {
                case 0: item.Id = (int)value!; break;
                case 1: item.Name = (string)value!; break;
                case 2: item.Price = (double)value!; break;
            }
        }
    }

    private static async Task<List<ExcelRow>> CollectAsync(IAsyncEnumerable<ExcelRow> rows)
    {
        var list = new List<ExcelRow>();
        await foreach (var row in rows) list.Add(row);
        return list;
    }

    private static void AssertSameRows(IReadOnlyList<ExcelRow> expected, IReadOnlyList<ExcelRow> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        for (var i = 0; i < expected.Count; i++)
        {
            Assert.Equal(expected[i].Index, actual[i].Index);
            Assert.Equal(expected[i].Values, actual[i].Values); // Typed compare: double vs double, DateTime vs DateTime...
        }
    }

    [Fact]
    public async Task ReadRowsAsync_MatchesSyncRowByRow()
    {
        var bytes = Workbook();
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var sync = reader.ReadRows("Products").ToList();
        var async = await CollectAsync(reader.ReadRowsAsync("Products"));

        AssertSameRows(sync, async);
        // Sanity: the data really covers every value kind the reader produces.
        var kinds = async.SelectMany(r => r.Values).Where(v => v != null).Select(v => v!.GetType()).Distinct().ToHashSet();
        Assert.Superset(new HashSet<Type> { typeof(string), typeof(double), typeof(bool), typeof(DateTime), typeof(TimeSpan) }, kinds);
        Assert.Contains(async, r => r.Values.Count == 0); // Blank rows survive as empty rows.
        Assert.Equal("Product 7 & co <x>", async.First(r => r.Index == 9).GetString(1)); // Row 7 is the blank row after product 5.
    }

    [Fact]
    public async Task ReadRowsAsync_ByIndex_AndCaseInsensitiveName()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook(3)));
        var byName = await CollectAsync(reader.ReadRowsAsync("PRODUCTS"));
        var byIndex = await CollectAsync(reader.ReadRowsAsync(0));
        AssertSameRows(byName, byIndex);

        var second = await CollectAsync(reader.ReadRowsAsync(1));
        Assert.Single(second);
        Assert.Equal(new object?[] { "a", 1.0 }, second[0].Values);
    }

    [Fact]
    public void ReadRowsAsync_FailsEagerlyForUnknownSheet()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook(1)));
        // Validation happens on the call, not on the first MoveNextAsync, like the sync API.
        Assert.Throws<ArgumentException>(() => reader.ReadRowsAsync("Nope"));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.ReadRowsAsync(5));
        Assert.Throws<ArgumentException>(() => reader.ReadAsync<Product>("Nope"));
        Assert.Throws<ArgumentException>(() => reader.ReadAsync("Nope", new ProductMapper()));
    }

    [Fact]
    public async Task ReadAsync_Reflection_MatchesSync()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook()));
        var sync = reader.Read<Product>("Products").ToList();
        var async = new List<Product>();
        await foreach (var p in reader.ReadAsync<Product>("Products")) async.Add(p);

        Assert.Equal(25, async.Count); // Blank rows and the sparse row (no bound column) are skipped.
        Assert.Equal(sync.Select(p => (p.Id, p.Name, p.Price, p.InStock, p.Shipped)), async.Select(p => (p.Id, p.Name, p.Price, p.InStock, p.Shipped)));
        Assert.Equal(new DateTime(2026, 1, 2), async[0].Shipped);
        Assert.Null(async[2].Shipped); // Row 3 was written empty.
    }

    [Fact]
    public async Task ReadAsync_WithHandWrittenMapper()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook(10)));
        var items = new List<Product>();
        await foreach (var p in reader.ReadAsync("Products", new ProductMapper())) items.Add(p);

        Assert.Equal(10, items.Count);
        Assert.Equal(Enumerable.Range(1, 10), items.Select(p => p.Id));
        Assert.Equal("Product 7 & co <x>", items[6].Name);
        Assert.Equal(12.5, items[9].Price);
        Assert.All(items, p => Assert.False(p.InStock)); // Not a mapper column: left at its default.

        var sync = reader.Read("Products", new ProductMapper()).Select(p => (p.Id, p.Name, p.Price));
        Assert.Equal(sync, items.Select(p => (p.Id, p.Name, p.Price)));
    }

    [Fact]
    public async Task ReadAsync_PrefersRegisteredMapper()
    {
        SheetMappers.Register(new ProductMapper());
        try
        {
            using var reader = ExcelReader.Open(new MemoryStream(Workbook(4)));
            var items = new List<Product>();
            await foreach (var p in reader.ReadAsync<Product>("Products")) items.Add(p);
            // The registered mapper does not bind InStock, so the reflection path was not taken.
            Assert.Equal(4, items.Count);
            Assert.All(items, p => Assert.False(p.InStock));
        }
        finally
        {
            SheetMappers.Unregister<Product>();
        }
    }

    [Fact]
    public async Task ReadDataTableAsync_MatchesSync()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook()));
        var sync = reader.ReadDataTable("Products");
        var async = await reader.ReadDataTableAsync("Products");

        Assert.Equal(sync.Columns.Cast<DataColumn>().Select(c => c.ColumnName), async.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        Assert.Equal(sync.Rows.Count, async.Rows.Count);
        for (var i = 0; i < sync.Rows.Count; i++) Assert.Equal(sync.Rows[i].ItemArray, async.Rows[i].ItemArray);
        Assert.Equal(7, async.Columns.Count); // Header has 5; the sparse row widened the table to G.

        var noHeader = await reader.ReadDataTableAsync("Products", firstRowIsHeader: false);
        Assert.Equal("A", noHeader.Columns[0].ColumnName);
        Assert.Equal(sync.Rows.Count + 1, noHeader.Rows.Count);
    }

    [Fact]
    public async Task OpenAsync_NonSeekableStream_IsBufferedAndClosed()
    {
        var bytes = Workbook();
        var source = new ForwardOnlyStream(new MemoryStream(bytes));
        using var reader = await ExcelReader.OpenAsync(source);

        Assert.True(source.Disposed); // leaveOpen defaults to false: the upload stream is released once copied.
        Assert.Equal(new[] { "Products", "Second" }, reader.SheetNames);
        using var expected = ExcelReader.Open(new MemoryStream(bytes));
        AssertSameRows(expected.ReadRows("Products").ToList(), await CollectAsync(reader.ReadRowsAsync("Products")));
    }

    [Fact]
    public async Task OpenAsync_LeaveOpen_KeepsTheSourceAlive()
    {
        var forwardOnly = new ForwardOnlyStream(new MemoryStream(Workbook(2)));
        using (var reader = await ExcelReader.OpenAsync(forwardOnly, leaveOpen: true))
            Assert.Equal(2, reader.SheetNames.Count);
        Assert.False(forwardOnly.Disposed);

        var seekable = new MemoryStream(Workbook(2));
        using (var reader = await ExcelReader.OpenAsync(seekable, leaveOpen: true))
            Assert.Equal(4, (await CollectAsync(reader.ReadRowsAsync(0))).Count); // Header, two products, the sparse row.
        Assert.True(seekable.CanRead); // Still open after the reader is disposed.
    }

    [Fact]
    public async Task OpenAsync_NonSeekable_SpillsToTempFileAboveTheLimit()
    {
        var bytes = Workbook(200);
        // The public limit is 64 MB; the internal entry point lets the test force the spill with a small one.
        using var buffered = await ExcelReader.BufferAsync(new ForwardOnlyStream(new MemoryStream(bytes)), memoryLimit: 4096, CancellationToken.None);

        var file = Assert.IsType<FileStream>(buffered);
        Assert.True(File.Exists(file.Name));
        Assert.Equal(0, buffered.Position);
        Assert.Equal(bytes.Length, buffered.Length);
        var copy = new byte[bytes.Length];
        Assert.Equal(bytes.Length, await buffered.ReadAtLeastAsync(copy, bytes.Length));
        Assert.Equal(bytes, copy);
        buffered.Position = 0;

        using (var reader = ExcelReader.Open(buffered, leaveOpen: true))
            Assert.Equal(200 + 1 + 40 + 1, (await CollectAsync(reader.ReadRowsAsync("Products"))).Count);
        buffered.Dispose();
        Assert.False(File.Exists(file.Name)); // DeleteOnClose: nothing is left behind.

        using var small = await ExcelReader.BufferAsync(new ForwardOnlyStream(new MemoryStream(bytes)), memoryLimit: long.MaxValue, CancellationToken.None);
        Assert.IsType<MemoryStream>(small);
        Assert.Equal(bytes.Length, small.Length);
    }

    [Fact]
    public async Task OpenAsync_Path_ReadsTheFile()
    {
        var path = Path.Combine(Path.GetTempPath(), "sheetstream-async-" + Guid.NewGuid().ToString("N") + ".xlsx");
        await File.WriteAllBytesAsync(path, Workbook(5));
        try
        {
            using var reader = await ExcelReader.OpenAsync(path);
            var rows = await CollectAsync(reader.ReadRowsAsync("Products"));
            Assert.Equal(5 + 1 + 1 + 1, rows.Count); // header, 5 products, one blank row (after 5), sparse row.
            Assert.Equal(1.0, rows[1][0]);
        }
        finally
        {
            File.Delete(path);
        }

        await Assert.ThrowsAsync<FileNotFoundException>(() => ExcelReader.OpenAsync(path)); // Reported through the task.
        await Assert.ThrowsAsync<ArgumentNullException>(async () => await ExcelReader.OpenAsync(default(string)!));
    }

    [Fact]
    public async Task OpenAsync_HonoursCancellationWhileBuffering()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelReader.OpenAsync(new ForwardOnlyStream(new MemoryStream(Workbook(1))), cancellationToken: cts.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExcelReader.OpenAsync("whatever.xlsx", cts.Token));
    }

    [Fact]
    public async Task ReadRowsAsync_CancellationMidSheet_ThrowsAndLeavesReaderUsable()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook(100)));
        using var cts = new CancellationTokenSource();
        var seen = 0;

        var ex = await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var row in reader.ReadRowsAsync("Products", cts.Token))
            {
                if (++seen == 10) cts.Cancel(); // The row in hand is delivered; the next MoveNextAsync throws.
            }
        });
        Assert.Equal(cts.Token, ex.CancellationToken);
        Assert.Equal(10, seen);

        // The iterator's finally ran: the entry stream is closed and the sheet can be read again in full.
        Assert.Equal(reader.ReadRows("Products").Count(), (await CollectAsync(reader.ReadRowsAsync("Products"))).Count);
        reader.Dispose(); // And disposing afterwards is clean (the outer using disposes a second time, which must also be harmless).
    }

    [Fact]
    public async Task ReadRowsAsync_WithCancellation_PreCancelledTokenThrowsBeforeTheFirstRow()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook(3)));
        using var cts = new CancellationTokenSource();
        cts.Cancel();

        // Token supplied through WithCancellation is combined with the one passed to the method.
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var _ in reader.ReadRowsAsync("Products").WithCancellation(cts.Token)) Assert.Fail("No row should be produced.");
        });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(async () =>
        {
            await foreach (var _ in reader.ReadAsync<Product>("Products", cts.Token)) Assert.Fail("No item should be produced.");
        });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadDataTableAsync("Products", cancellationToken: cts.Token));
    }

    [Fact]
    public async Task ReadRowsAsync_BreakingEarly_ReleasesTheSheet()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Workbook(50)));
        var firstTwo = new List<int>();
        await foreach (var row in reader.ReadRowsAsync("Products"))
        {
            firstTwo.Add(row.Index);
            if (firstTwo.Count == 2) break;
        }
        Assert.Equal(new[] { 1, 2 }, firstTwo);

        // Two interleaved enumerations and a full re-read all work: nothing was left half-open.
        await using var a = reader.ReadRowsAsync("Products").GetAsyncEnumerator();
        await using var b = reader.ReadRowsAsync("Second").GetAsyncEnumerator();
        Assert.True(await a.MoveNextAsync());
        Assert.True(await b.MoveNextAsync());
        Assert.Equal("a", b.Current.GetString(0));
        Assert.Equal(reader.ReadRows("Products").Count(), (await CollectAsync(reader.ReadRowsAsync("Products"))).Count);
    }

    [Fact]
    public async Task EndToEnd_UploadRoundTrip()
    {
        // The scenario the API exists for: a request body arrives as a non-seekable stream and is mapped to objects.
        var products = Enumerable.Range(1, 300).Select(i => new Product { Id = i, Name = "P" + i, Price = i * 0.5, InStock = i % 2 == 0, Shipped = i % 4 == 0 ? null : new DateTime(2026, 3, 1).AddHours(i) }).ToList();
        var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Products").LoadFromCollection(products));

        using var reader = await ExcelReader.OpenAsync(new ForwardOnlyStream(new MemoryStream(bytes)));
        var back = new List<Product>();
        await foreach (var p in reader.ReadAsync<Product>("Products").WithCancellation(CancellationToken.None)) back.Add(p);

        Assert.Equal(products.Select(p => (p.Id, p.Name, p.Price, p.InStock, p.Shipped)), back.Select(p => (p.Id, p.Name, p.Price, p.InStock, p.Shipped)));
    }
}
