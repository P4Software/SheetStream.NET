using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using SheetStream.Internal;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// The row writer formats numbers into a reusable buffer instead of asking the XML writer to build a string for
/// each one. These tests pin the two promises that makes: the text is exactly what the string-building call
/// produced (so no file changes), and writing a cell allocates nothing.
/// </summary>
public class CellWritingAllocationTests
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    /// <summary>Numbers that have burned number formatting before: zero, negative zero, exponents, extremes.</summary>
    private static readonly double[] AwkwardDoubles =
    {
        0, -0.0, 1, -1, 0.1, -0.1, 1.0 / 3, 2.0 / 3, 100, 1e15, 1e16, 123456789012345678, 1e21, 1e-5, 1.5e-7, 12345.6789,
        0.30000000000000004, double.Epsilon, -double.Epsilon, double.MaxValue, double.MinValue, 4.9e-324, 1.7976931348623157e308,
        9007199254740993, 45000.5, 45000.999988425923, 36526, 0.000001, 1234567.891,
    };

    private static readonly long[] AwkwardLongs =
    {
        0, 1, -1, 9, 10, -10, 99, 100, 299, 300, 301, int.MaxValue, int.MinValue, (long)int.MaxValue + 1, (long)int.MinValue - 1,
        long.MaxValue, long.MinValue, long.MinValue + 1, 999999999999999999, 1000000000000000000,
    };

    private static readonly decimal[] AwkwardDecimals =
    {
        0m, 1m, -1m, 1.10m, 0.0000000000000000000000000001m, decimal.MaxValue, decimal.MinValue, 123456789.123456789m, -0.5m, 100m, 79228162514264337593543950335m,
    };

    private static string Text(byte[] buffer, int length) => System.Text.Encoding.ASCII.GetString(buffer, 0, length);

    [Fact]
    public void Integers_FormatLikeTheInvariantCulture()
    {
        var buffer = new byte[Utf8Numbers.MaxLength];
        foreach (var value in AwkwardLongs)
            Assert.Equal(XmlConvert.ToString(value), Text(buffer, Utf8Numbers.Format(value, buffer, 0)));

        var random = new Random(20261115); // Fixed seed: a failure must reproduce.
        for (var i = 0; i < 200_000; i++)
        {
            // Spread over every magnitude, not only the huge ones a raw NextInt64 would give.
            var value = random.NextInt64() >> random.Next(0, 64);
            if (random.Next(2) == 0) value = -value;
            Assert.Equal(XmlConvert.ToString(value), Text(buffer, Utf8Numbers.Format(value, buffer, 0)));
        }
    }

    [Fact]
    public void Doubles_FormatLikeXmlConvert()
    {
        var buffer = new byte[Utf8Numbers.MaxLength];
        foreach (var value in AwkwardDoubles)
            Assert.Equal(XmlConvert.ToString(value), Text(buffer, Utf8Numbers.Format(value, buffer, 0)));

        var random = new Random(20261116);
        for (var i = 0; i < 200_000; i++)
        {
            // Half from raw bit patterns (every exponent), half from the everyday range of money and dates.
            var value = i % 2 == 0
                ? BitConverter.Int64BitsToDouble(random.NextInt64())
                : Math.Round((random.NextDouble() - 0.3) * Math.Pow(10, random.Next(-4, 9)), random.Next(0, 6));
            if (double.IsNaN(value) || double.IsInfinity(value)) continue; // Never reaches the formatter: the writer turns them into #NUM!.
            Assert.Equal(XmlConvert.ToString(value), Text(buffer, Utf8Numbers.Format(value, buffer, 0)));
        }
    }

    /// <summary>
    /// The fast path for short decimals (<c>1234.5</c>, <c>0.07</c>) must give the text of the runtime's own shortest
    /// round-trip formatter for every value it accepts, and leave the rest to it. These shapes are the ones the fast
    /// path decides on: decimals of every length and magnitude, the near misses (binary noise like 0.1 + 0.2) and the
    /// edges of its range. A fixed seed, so a failure reproduces.
    /// </summary>
    [Fact]
    public void Doubles_FastPathAgreesWithTheRuntimeFormatter()
    {
        var buffer = new byte[Utf8Numbers.MaxLength];
        void Check(double value)
        {
            if (double.IsNaN(value) || double.IsInfinity(value)) return;
            var written = Utf8Numbers.Format(value, buffer, 3);
            Assert.Equal(value.ToString("R", System.Globalization.CultureInfo.InvariantCulture), System.Text.Encoding.ASCII.GetString(buffer, 3, written));
        }

        foreach (var edge in new[] { 0.0, -0.0, 1e-4, 9.999999999999999e-5, 0.0001, 0.00011, 1e-5, 1, 1e14, 123456789012345.0, 999999999999999.0, 999999999999999.9, 1e15, 4503599627370496.0, 0.1 + 0.2, 1.0000000000000002 })
        {
            Check(edge);
            Check(-edge);
        }

        var random = new Random(20261118);
        for (var i = 0; i < 1_500_000; i++)
        {
            var places = random.Next(0, 9);
            var digits = random.Next(1, 16);
            var value = (long)(random.NextDouble() * Math.Pow(10, digits)) / Math.Pow(10, places);
            Check(value);
            Check(-value);
            Check(value + random.Next(1, 100) / 100.0);            // Sums of short decimals: often noise.
            Check(value * 1.1);
            Check(random.Next(0, 10000) * 1.25 + 0.99);
            Check(Math.Round(random.NextDouble() * 1e6, random.Next(0, 7)));
            Check(random.NextDouble() * Math.Pow(10, random.Next(-8, 18)));
        }
    }

    [Fact]
    public void Decimals_FormatLikeXmlConvert()
    {
        var buffer = new byte[Utf8Numbers.MaxLength];
        foreach (var value in AwkwardDecimals)
            Assert.Equal(XmlConvert.ToString(value), Text(buffer, Utf8Numbers.Format(value, buffer, 0)));

        var random = new Random(20261117);
        for (var i = 0; i < 50_000; i++)
        {
            var value = new decimal(random.Next(), random.Next(), random.Next(0, 1 << 20), random.Next(2) == 0, (byte)random.Next(0, 29));
            Assert.Equal(XmlConvert.ToString(value), Text(buffer, Utf8Numbers.Format(value, buffer, 0)));
        }
    }

    /// <summary>
    /// The same promise at the level of the file: whatever a cell holds, the text inside its <c>v</c> element is
    /// what the string-building call used to write, including style ids and shared string indexes past 299
    /// (the largest numbers .NET keeps as ready-made strings).
    /// </summary>
    [Fact]
    public void WrittenSheet_HoldsTheSameNumberText()
    {
        var doubles = AwkwardDoubles.Where(d => !double.IsInfinity(d)).ToArray();
        SheetStream.Styles.CellStyle[] styles = null!;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            styles = Enumerable.Range(0, 400).Select(i => wb.Styles.Create(o => o.Fill.BackgroundColor = $"#{i:X6}")).ToArray();
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell(doubles[0]).WriteCell(float.MaxValue).WriteCell(0.1f); // The float overload widens to double.
            foreach (var d in doubles.Skip(1)) s.Row().WriteCell(d, styles[(int)(Math.Abs(d) % 400)]);
            foreach (var l in AwkwardLongs) s.Row().WriteCell(l);
            foreach (var l in AwkwardLongs) s.Row().WriteCell((int)Math.Clamp(l, int.MinValue, int.MaxValue));
            foreach (var m in AwkwardDecimals) s.Row().WriteCell(m);
            foreach (var style in styles) s.Row().WriteCell(1, style);                         // Style ids 1..400.
            for (var i = 0; i < 400; i++) s.Row().WriteCell("string " + i);                    // Shared string indexes 0..399.
            s.Row().WriteCell(double.NaN).WriteCell(double.PositiveInfinity);                  // Error cells, not numbers.
        }, new PackageOptions { StringStorage = StringStorage.Shared }); // Text goes through the table, so its indexes are in the sheet.

        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        XDocument sheet;
        using (var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()) sheet = XDocument.Load(stream);
        var values = sheet.Descendants(Main + "c").Select(c => (string?)c.Element(Main + "v")).ToList();

        var expected = new List<string?>
        {
            XmlConvert.ToString(doubles[0]), XmlConvert.ToString((double)float.MaxValue), XmlConvert.ToString((double)0.1f),
        };
        expected.AddRange(doubles.Skip(1).Select(XmlConvert.ToString));
        expected.AddRange(AwkwardLongs.Select(XmlConvert.ToString));
        expected.AddRange(AwkwardLongs.Select(l => XmlConvert.ToString((int)Math.Clamp(l, int.MinValue, int.MaxValue))));
        expected.AddRange(AwkwardDecimals.Select(XmlConvert.ToString));
        expected.AddRange(Enumerable.Repeat("1", 400));
        expected.AddRange(Enumerable.Range(0, 400).Select(i => i.ToString()));
        expected.AddRange(new[] { "#NUM!", "#NUM!" });
        Assert.Equal(expected, values);

        // Style ids come back out of the same writer: every one of the 400 styles is referenced by its own id.
        var styleIds = sheet.Descendants(Main + "c").Where(c => c.Element(Main + "v")?.Value == "1" && c.Attribute("s") != null)
            .Select(c => (int)c.Attribute("s")!).ToList();
        Assert.Equal(styles.Select(s => s.Id), styleIds.Skip(styleIds.Count - 400));

        // And the file still reads back to the numbers that went in.
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("S").ToList();
        Assert.Equal(AwkwardLongs.Select(l => (double)l), rows.Skip(doubles.Length).Take(AwkwardLongs.Length).Select(r => (double)r[0]!));
    }

    /// <summary>
    /// Writing a cell allocates nothing. Measured as a difference between a short and a long run of the same
    /// shape in one process, so the one-off costs (compression buffers, the style sheet, the string table for the
    /// repeated values) cancel out, and what is left is the price of the extra cells.
    /// </summary>
    [Fact]
    public void WritingCells_AllocatesNothingPerCell()
    {
        // Warm everything once (JIT, caches, pooled buffers) so neither measured run pays for it.
        WriteSheet(2_000);
        var small = Measure(2_000);
        var large = Measure(22_000);

        const int ExtraCells = 20_000 * 10;
        var perCell = (large - small) / (double)ExtraCells;
        // Before the change this was about 25 bytes per cell (the strings made for every number).
        Assert.True(perCell < 0.5, $"Writing a cell allocated {perCell:F2} bytes on average ({large - small:N0} bytes over {ExtraCells:N0} extra cells).");
    }

    /// <summary>
    /// The cell kinds the first test leaves out and the characters the XML writer escapes: text with <c>&amp;</c>,
    /// <c>&lt;</c> and <c>&gt;</c> (escaped by hand into the staging buffer), text with surrounding spaces, a long
    /// text that is written in pieces, gaps that need the cell address, empty styled cells, floats and durations.
    /// </summary>
    [Fact]
    public void WritingEscapedTextAndGaps_AllocatesNothingPerCell()
    {
        WriteSheet(2_000, extended: true);
        var small = Measure(2_000, extended: true);
        var large = Measure(22_000, extended: true);

        const int ExtraCells = 20_000 * 8; // Eight cells per row written; the Skip is not a cell.
        var perCell = (large - small) / (double)ExtraCells;
        Assert.True(perCell < 0.5, $"Writing a cell allocated {perCell:F2} bytes on average ({large - small:N0} bytes over {ExtraCells:N0} extra cells).");
    }

    private static long Measure(int rows, bool extended = false)
    {
        var before = GC.GetAllocatedBytesForCurrentThread();
        WriteSheet(rows, extended);
        return GC.GetAllocatedBytesForCurrentThread() - before;
    }

    // A cell holds at most 32,767 characters, so this text (32,003 characters, 32 KB or more as UTF-8) is longer than the write buffer
    // only because WriteSheet sets the buffer to LongTextBuffer, not the 64 KB default.
    private const int LongTextBuffer = 16 * 1024;

    private static readonly string LongText = new string('x', 16_000) + " & " + new string('y', 16_000);

    private static void WriteSheet(int rows, bool extended = false)
    {
        var names = new[] { "North", "South", "East", "West" };
        var date = new DateTime(2026, 1, 1);
        using var package = new ExcelPackage(new DiscardingStream(), leaveOpen: true, new PackageOptions { CompressionLevel = CompressionLevel.Fastest, BufferSize = LongTextBuffer });
        var money = package.Workbook.Styles.Create(o => o.NumberFormat = "#,##0.00");
        var dated = package.Workbook.Styles.Create(o => o.NumberFormat = "yyyy-mm-dd");
        var sheet = package.Workbook.Worksheets.Add("S");
        for (var i = 0; i < rows; i++)
        {
            if (extended)
            {
                sheet.Row()
                    .WriteInlineString("Tom & Jerry <fun> 5 > 3")                // Entities in the text.
                    .WriteInlineString("  padded  ")                           // xml:space="preserve".
                    .WriteCell("a & b")                                        // A shared string with an entity.
                    .Skip(2).WriteCell(i, dated)                               // A gap, so this cell carries its address.
                    .WriteEmpty(money)                                         // A styled cell with no value.
                    .WriteCell((float)i / 3)                                   // float
                    .WriteCell(TimeSpan.FromMinutes(i))                        // duration
                    .WriteInlineString(i % 1000 == 0 ? LongText : "short");    // Now and then a text longer than the (16 KB) buffer.
                continue;
            }
            sheet.Row()
                .WriteCell(i)                         // int
                .WriteCell(names[i & 3])              // shared string
                .WriteCell(names[(i >> 2) & 3])       // another
                .WriteCell((long)i * 100_000)         // long
                .WriteCell(i * 1.25 + 0.01, money)    // double with a style
                .WriteCell(i / 7.0, money)            // a long double
                .WriteCell(date.AddDays(i % 365), dated)
                .WriteCell((i & 1) == 0)              // bool
                .WriteInlineString(names[i & 3])      // inline string
                .WriteCell(i * 0.5m);                 // decimal
        }
        package.Save();
    }

    /// <summary>A write-only stream that keeps nothing, so growing output buffers do not show up as allocations.</summary>
    private sealed class DiscardingStream : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) { }
    }
}
