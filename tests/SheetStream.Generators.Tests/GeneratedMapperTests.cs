using System.IO.Compression;
using SheetStream.Reading;

namespace SheetStream.Generators.Tests;

/// <summary>End-to-end checks of the generated mappers against the reflection path they replace.</summary>
public class GeneratedMapperTests
{
    private static readonly string[] ExpectedHeaders =
    {
        // Order 0, Order 1 (field), Order 5 (field), then the unordered members: fields first (metadata token
        // order, as reflection does), then properties in declaration order.
        "Order #", "Flags", "Reference", "Customer name", "Amount", "Placed", "Paid", "State", "Discount", "Key", "Duration", "Big", "Small", "Ratio", "Shipped",
    };

    private static List<Order> SampleOrders() => Enumerable.Range(1, 30).Select(i => new Order
    {
        Id = i,
        Customer = i % 7 == 0 ? null : "Customer " + i,
        Amount = i * 10.25m,
        Placed = new DateTime(2026, 1, i % 28 + 1, i % 24, 0, 0),
        Paid = i % 2 == 0,
        State = (Status)(i % 3),
        Discount = i % 4 == 0 ? null : i / 100.0,
        Reference = "REF-" + i,
        Key = new Guid(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0),
        Duration = TimeSpan.FromMinutes(i * 7),
        Big = 10_000_000_000L + i,
        Small = (short)(-i),
        Ratio = i / 3f,
        Shipped = i % 5 == 0 ? null : new DateTimeOffset(2026, 2, 1, 12, 0, 0, TimeSpan.Zero).AddDays(i),
        Flags = (byte)(i % 256),
    }).ToList();

    private static OrderPlain ToPlain(Order o) => new()
    {
        Id = o.Id, Customer = o.Customer, Amount = o.Amount, Placed = o.Placed, Paid = o.Paid, State = o.State,
        Discount = o.Discount, Reference = o.Reference, Key = o.Key, Duration = o.Duration, Big = o.Big,
        Small = o.Small, Ratio = o.Ratio, Shipped = o.Shipped, Flags = o.Flags,
    };

    [Fact]
    public void Mappers_AreRegisteredAutomatically()
    {
        Assert.True(SheetMappers.TryGet<Order>(out var order));
        Assert.Same(Order.SheetMapper, order);
        Assert.True(SheetMappers.TryGet<Geometry.Shapes.Point>(out var point));
        Assert.Same(Geometry.Shapes.Point.SheetMapper, point);
        Assert.True(SheetMappers.TryGet<Line>(out _));
        Assert.False(SheetMappers.TryGet<OrderPlain>(out _));
        Assert.True(Holder.HiddenMapperWorks());
    }

    [Fact]
    public void GeneratedMapper_DescribesColumnsLikeReflection()
    {
        var m = Order.SheetMapper;
        Assert.Equal(ExpectedHeaders, m.Headers);
        Assert.Equal("#,##0.00", m.NumberFormats[4]);
        Assert.All(m.NumberFormats.Where((_, i) => i != 4), f => Assert.Null(f));
        Assert.True(m.CanRead);
        Assert.Equal(typeof(int), m.MemberType(0));
        Assert.Equal(typeof(double?), m.MemberType(8));
        Assert.Equal(typeof(string), m.MemberType(3));
        Assert.False(m.CanSet(9)); // init-only Key
        Assert.True(m.CanSet(2));  // public field
        Assert.Equal(0, m.FindColumn("order #"));
        Assert.Equal(0, m.FindColumn("ID")); // member name, case-insensitive
        Assert.Equal(3, m.FindColumn("Customer name"));
        Assert.Equal(3, m.FindColumn("Customer"));
        Assert.Equal(-1, m.FindColumn("Key"));    // not settable
        Assert.Equal(-1, m.FindColumn("Secret")); // ignored
        Assert.Equal(-1, m.FindColumn("nope"));

        var line = Line.SheetMapper;
        Assert.Equal(new[] { "Qty", "Sku" }, line.Headers);
        Assert.False(line.CanRead);
        Assert.Throws<NotSupportedException>(() => line.Create());

        var point = Geometry.Shapes.Point.SheetMapper;
        Assert.Equal(new[] { "X", "Y", "Label" }, point.Headers);
        var p = point.Create();
        point.SetValue(ref p, 0, 3);
        point.SetValue(ref p, 2, "origin");
        Assert.Equal((3, "origin"), (p.X, p.Name));
        Assert.Equal(3, point.GetValue(p, 0));
    }

    [Fact]
    public void GeneratedWriter_ProducesTheSameWorkbookAsReflection()
    {
        var orders = SampleOrders();
        var viaMapper = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Orders").LoadFromCollection(orders, Order.SheetMapper, new LoadOptions { TableName = "T" }));
        var viaRegistry = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Orders").LoadFromCollection(orders, new LoadOptions { TableName = "T" }));
        var viaReflection = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Orders").LoadFromCollection(orders.Select(ToPlain).ToList(), new LoadOptions { TableName = "T" }));

        // The sheet, the styles (number formats) and the shared strings must be byte-identical: same headers,
        // same values, same formats, same cell types.
        foreach (var part in new[] { "xl/worksheets/sheet1.xml", "xl/styles.xml", "xl/sharedStrings.xml", "xl/tables/table1.xml" })
        {
            var expected = ReadPart(viaReflection, part);
            Assert.Equal(expected, ReadPart(viaMapper, part));
            Assert.Equal(expected, ReadPart(viaRegistry, part));
        }

        using var a = ExcelReader.Open(new MemoryStream(viaMapper));
        using var b = ExcelReader.Open(new MemoryStream(viaReflection));
        var rowsA = a.ReadRows("Orders").ToList();
        var rowsB = b.ReadRows("Orders").ToList();
        Assert.Equal(31, rowsA.Count);
        Assert.Equal(ExpectedHeaders, rowsA[0].Values.Select(v => (string?)v));
        for (var i = 0; i < rowsA.Count; i++) Assert.Equal(rowsB[i].Values, rowsA[i].Values);
    }

    [Fact]
    public void GeneratedReader_RoundTripsThroughTheMapper()
    {
        var orders = SampleOrders();
        var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Orders").LoadFromCollection(orders));

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var viaRegistry = reader.Read<Order>("Orders").ToList();
        var viaExplicit = reader.Read("Orders", Order.SheetMapper).ToList();
        var viaReflection = reader.Read<OrderPlain>("Orders").ToList();
        Assert.Equal(orders.Count, viaRegistry.Count);
        Assert.Equal(orders.Count, viaExplicit.Count);
        Assert.Equal(orders.Count, viaReflection.Count);

        for (var i = 0; i < orders.Count; i++)
        {
            foreach (var back in new[] { viaRegistry[i], viaExplicit[i] })
            {
                var o = orders[i];
                Assert.Equal(o.Id, back.Id);
                Assert.Equal(o.Customer, back.Customer);
                Assert.Equal(o.Amount, back.Amount);
                Assert.Equal(o.Placed, back.Placed);
                Assert.Equal(o.Paid, back.Paid);
                Assert.Equal(o.State, back.State);
                Assert.Equal(o.Discount, back.Discount);
                Assert.Equal(o.Reference, back.Reference);
                // Durations travel as a fraction of a day in a double, so both paths lose a tick; they must agree.
                Assert.Equal(viaReflection[i].Duration, back.Duration);
                Assert.InRange((o.Duration - back.Duration).Duration(), TimeSpan.Zero, TimeSpan.FromMilliseconds(1));
                Assert.Equal(o.Big, back.Big);
                Assert.Equal(o.Small, back.Small);
                Assert.Equal(o.Ratio, back.Ratio);
                Assert.Equal(o.Shipped?.DateTime, back.Shipped?.DateTime);
                Assert.Equal(o.Flags, back.Flags);
                Assert.Equal("hidden", back.Secret);
                // Init-only members cannot be assigned from generated code; reflection still sets them.
                Assert.Equal(Guid.Empty, back.Key);
            }
            Assert.Equal(orders[i].Key, viaReflection[i].Key);
            Assert.Equal(orders[i].Big, viaReflection[i].Big);
        }
    }

    [Fact]
    public void GeneratedReader_ReportsConversionErrorsWithRowAndColumn()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteCell("Order #").WriteCell("Amount");
            s.Row().WriteCell("not a number").WriteCell(1.5);
        });
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var ex = Assert.Throws<FormatException>(() => reader.Read("S", Order.SheetMapper).ToList());
        Assert.Contains("Row 2, column A ('Order #')", ex.Message);
        Assert.Throws<NotSupportedException>(() => reader.Read("S", Line.SheetMapper));
    }

    [Fact]
    public void StructMapper_WritesAndReadsNestedStruct()
    {
        var points = new[] { new Geometry.Shapes.Point { X = 1, Y = 2, Name = "a" }, new Geometry.Shapes.Point { X = 3, Y = 4 } };
        var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("P").LoadFromCollection(points));
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var back = reader.Read<Geometry.Shapes.Point>("P").ToList();
        Assert.Equal(2, back.Count);
        Assert.Equal((1, 2, "a"), (back[0].X, back[0].Y, back[0].Name));
        Assert.Equal((3, 4, (string?)null), (back[1].X, back[1].Y, back[1].Name));
    }

    [Fact]
    public async Task AsyncEnumerable_UsesTheMapper()
    {
        var orders = SampleOrders();
        static async IAsyncEnumerable<T> Stream<T>(IEnumerable<T> items)
        {
            foreach (var item in items) { await Task.Yield(); yield return item; }
        }

        using var ms = new MemoryStream();
        using (var package = new ExcelPackage(ms, leaveOpen: true))
        {
            var sheet = package.Workbook.Worksheets.Add("A");
            var result = await sheet.LoadFromAsyncEnumerable(Stream(orders), Order.SheetMapper, new LoadOptions { TableName = "T" });
            Assert.Equal(orders.Count, result.DataRows);
            package.Save();
        }
        var sync = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("A").LoadFromCollection(orders, new LoadOptions { TableName = "T" }));
        Assert.Equal(ReadPart(sync, "xl/worksheets/sheet1.xml"), ReadPart(ms.ToArray(), "xl/worksheets/sheet1.xml"));
    }

    private static string ReadPart(byte[] package, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        var entry = zip.GetEntry(name) ?? throw new InvalidOperationException("Missing " + name);
        using var sr = new StreamReader(entry.Open());
        return sr.ReadToEnd();
    }
}
