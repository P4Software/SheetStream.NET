using System.Text;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream.Tests;

/// <summary>
/// Tests for the reader that sits between a package part and the XML parser. Two things are checked: that it
/// never changes the text (a differential test against plain decoding, with awkward read sizes), and that its
/// limits hold wherever a marker or a limit falls relative to its internal buffer.
/// </summary>
public class GuardedPartReaderTests
{
    // The reader decodes 16 KB of bytes at a time; tests place markers on both sides of that edge.
    private const int Edge = 16 * 1024;

    /// <summary>Hands out the bytes in pieces of a chosen or random size, like a deflate stream may.</summary>
    private sealed class ChunkStream : MemoryStream
    {
        private readonly Random? _random;
        private readonly int _fixed;

        public ChunkStream(byte[] bytes, int chunk, Random? random = null) : base(bytes, writable: false)
        {
            _fixed = chunk;
            _random = random;
        }

        private int Next(int count) => Math.Min(count, _random != null ? _random.Next(1, 4000) : _fixed);

        public override int Read(byte[] buffer, int offset, int count) => base.Read(buffer, offset, Next(count));

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            => base.ReadAsync(buffer, offset, Next(count), cancellationToken);
    }

    private static string ReadAll(byte[] bytes, long maxBytes = long.MaxValue, long maxRun = long.MaxValue, int chunk = int.MaxValue, Random? random = null, long maxTag = long.MaxValue)
    {
        using var reader = new GuardedPartReader(new ChunkStream(bytes, chunk, random), "part.xml", maxBytes, true, maxRun, maxTag);
        var sb = new StringBuilder();
        var buffer = new char[5000];
        int n;
        while ((n = reader.Read(buffer, 0, random?.Next(1, buffer.Length) ?? buffer.Length)) > 0) sb.Append(buffer, 0, n);
        return sb.ToString();
    }

    private static async Task<string> ReadAllAsync(byte[] bytes, long maxBytes, long maxRun, Random random)
    {
        using var reader = new GuardedPartReader(new ChunkStream(bytes, 0, random), "part.xml", maxBytes, true, maxRun, long.MaxValue);
        var sb = new StringBuilder();
        var buffer = new char[5000];
        int n;
        while ((n = await reader.ReadAsync(buffer, 0, random.Next(1, buffer.Length))) > 0) sb.Append(buffer, 0, n);
        return sb.ToString();
    }

    private static byte[] Utf8(string text) => new UTF8Encoding(false).GetBytes(text);

    /// <summary>Text built from the pieces that matter to the scanner, in random order.</summary>
    private static string RandomDocument(Random random, int pieces)
    {
        string[] parts =
        {
            "<row r=\"1\">", "</row>", "<c><v>12.5</v></c>", "plain text ", "!", "<!", "<![", "<![CDATA[", "]]>", "]", "]]", "]]]>",
            "<![CDATA[ a < b ]]>", "Sheet1!A1", "é", "日本語", "\U0001F600", "&amp;", "<", ">", "[CDATA[", new string('x', 300), new string(']', 40), "\r\n",
        };
        var sb = new StringBuilder();
        for (var i = 0; i < pieces; i++) sb.Append(parts[random.Next(parts.Length)]);
        return sb.ToString();
    }

    [Fact]
    public async Task Text_IsPassedThroughUnchanged_WhateverTheReadSizes()
    {
        for (var seed = 1; seed <= 60; seed++)
        {
            var random = new Random(seed);
            var text = RandomDocument(random, random.Next(1, 4000));
            var bytes = Utf8(text);

            Assert.Equal(text, ReadAll(bytes, random: new Random(seed * 7)));
            Assert.Equal(text, await ReadAllAsync(bytes, long.MaxValue, long.MaxValue, new Random(seed * 11)));
        }
    }

    [Fact]
    public void Text_IsPassedThroughUnchanged_WhenBytesArriveOneAtATime()
    {
        var text = RandomDocument(new Random(5), 800) + "<![CDATA[ tail ]]>";
        Assert.Equal(text, ReadAll(Utf8(text), chunk: 1));
        Assert.Equal("", ReadAll(Array.Empty<byte>()));
        Assert.Equal("<a/>", ReadAll(Utf8("<a/>"), chunk: 1));
    }

    [Theory]
    [InlineData("utf-8", false)]
    [InlineData("utf-8", true)]
    [InlineData("utf-16", true)]
    [InlineData("utf-16BE", true)]
    [InlineData("utf-16", false)]
    [InlineData("utf-16BE", false)]
    [InlineData("utf-32", true)]
    [InlineData("utf-32BE", true)]
    public void EveryUnicodeEncoding_IsDecoded(string name, bool byteOrderMark)
    {
        var text = "<?xml version=\"1.0\"?><sst><si><t>añ日本\U0001F600 " + new string('é', 40_000) + "</t></si></sst>";
        var encoding = Encoding.GetEncoding(name);
        var body = encoding.GetBytes(text);
        var bytes = byteOrderMark ? encoding.GetPreamble().Concat(body).ToArray() : body;

        Assert.Equal(text, ReadAll(bytes));
        Assert.Equal(text, ReadAll(bytes, chunk: 3));
    }

    [Fact]
    public void EncodingNamedInTheDeclaration_IsUsed()
    {
        var text = "<?xml version='1.0' encoding='ISO-8859-1'?><t>café ñ</t>";
        Assert.Equal(text, ReadAll(Encoding.Latin1.GetBytes(text)));
    }

    [Fact]
    public void BytesThatAreNotValidText_AreRejectedAsDamage()
    {
        var bytes = Utf8("<t>ok").Concat(new byte[] { 0xC3, 0x28, 0xFF }).Concat(Utf8("</t>")).ToArray();
        Assert.Throws<InvalidDataException>(() => ReadAll(bytes));
        // A lone surrogate half in UTF-16.
        var utf16 = new byte[] { 0xFF, 0xFE, (byte)'<', 0, (byte)'t', 0, (byte)'>', 0, 0x00, 0xD8, (byte)'<', 0 };
        Assert.Throws<InvalidDataException>(() => ReadAll(utf16));
    }

    [Theory]
    [InlineData("no-such-encoding")]
    [InlineData("utf-16")]   // Named in a part whose bytes are plainly not UTF-16.
    public void DeclarationNamingAnUnusableEncoding_IsRejected(string name)
    {
        var bytes = Utf8($"<?xml version=\"1.0\" encoding=\"{name}\"?><t>x</t>");
        Assert.Throws<InvalidDataException>(() => ReadAll(bytes));
    }

    [Fact]
    public async Task RefusedPart_StaysRefused_WhenReadAgain()
    {
        var bytes = Utf8("<a>" + string.Concat(Enumerable.Repeat("<b/>", 20_000)) + new string('x', 5000) + "</a>" + string.Concat(Enumerable.Repeat("<b/>", 20_000)));
        using var reader = new GuardedPartReader(new MemoryStream(bytes), "part.xml", long.MaxValue, true, 1000, long.MaxValue);
        var buffer = new char[4096];
        var first = Assert.Throws<InvalidDataException>(() => { while (reader.Read(buffer, 0, buffer.Length) > 0) { } });
        for (var i = 0; i < 50; i++)
        {
            Assert.Equal(first.Message, Assert.Throws<InvalidDataException>(() => reader.Read(buffer, 0, buffer.Length)).Message);
            Assert.Equal(first.Message, (await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(buffer, 0, buffer.Length))).Message);
        }
    }

    [Fact]
    public void TotalByteLimit_IsExact()
    {
        var bytes = Utf8("<a>" + new string('x', 50_000) + "</a>");
        Assert.Equal(bytes.Length, ReadAll(bytes, maxBytes: bytes.Length).Length);
        var ex = Assert.Throws<InvalidDataException>(() => ReadAll(bytes, maxBytes: bytes.Length - 1));
        Assert.Contains("MaxLoadedPartBytes", ex.Message);
    }

    [Theory]
    [InlineData(100)]
    [InlineData(70_000)]   // Longer than the internal buffer, so the run spans several reads.
    public void RunLimit_IsExact_ForTextAndForATag(int limit)
    {
        // limit characters between two '<' is allowed; one more is not. The tag form counts its own "a x=" etc.
        var atLimit = "<a>" + new string('x', limit - 2) + "</a>";       // "a>" + text = limit characters after '<'
        var over = "<a>" + new string('x', limit - 1) + "</a>";
        Assert.Equal(atLimit, ReadAll(Utf8(atLimit), maxRun: limit, chunk: 977));
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8(over), maxRun: limit, chunk: 977));

        // After the '<' come a x=" (5 characters), the value, and "/> (3): limit characters in all.
        var tagAtLimit = "<a x=\"" + new string('q', limit - 8) + "\"/><b/>";
        var tagOver = "<a x=\"" + new string('q', limit - 7) + "\"/><b/>";
        Assert.Equal(tagAtLimit, ReadAll(Utf8(tagAtLimit), maxRun: limit));
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8(tagOver), maxRun: limit));
    }

    [Fact]
    public void TextAfterTheLastTag_CountsToo()
    {
        var bytes = Utf8("<a/>" + new string(' ', 5000));
        Assert.Throws<InvalidDataException>(() => ReadAll(bytes, maxRun: 1000));
        // Exact: "a/>" plus 997 spaces is 1,000 characters after the last '<'.
        Assert.Equal(1001, ReadAll(Utf8("<a/>" + new string(' ', 997)), maxRun: 1000).Length);
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8("<a/>" + new string(' ', 998)), maxRun: 1000));
    }

    [Fact]
    public void LongCData_IsRejected_WhereverItsMarkerFallsInTheBuffer()
    {
        const int limit = 1000;
        var hostile = "<![CDATA[" + new string('<', limit + 50) + "]]>";
        var genuine = "<![CDATA[" + new string('<', limit - 50) + "]]>";

        // Every padding that puts the marker before, across and after the first buffer edge, and the second.
        var paddings = Enumerable.Range(Edge - 40, 80).Concat(Enumerable.Range(2 * Edge - 40, 80)).Concat(new[] { 0, 1, 7, 8, 9 });
        foreach (var padding in paddings)
        {
            // The padding is short tags, so it never trips the run limit itself.
            var lead = string.Concat(Enumerable.Repeat("<p/>", padding / 4)) + new string(' ', padding % 4);
            foreach (var chunk in new[] { int.MaxValue, 1000, 13 })
            {
                var ok = "<a>" + lead + genuine + "</a>";
                Assert.Equal(ok, ReadAll(Utf8(ok), maxRun: limit, chunk: chunk));
                var bad = Utf8("<a>" + lead + hostile + "</a>");
                Assert.Throws<InvalidDataException>(() => ReadAll(bad, maxRun: limit, chunk: chunk));
            }
        }
    }

    [Fact]
    public void CDataEnd_IsFound_WhereverItFallsInTheBuffer()
    {
        const int limit = 1000;
        // After the section ends the count starts again, so long genuine markup after it must pass; if the
        // end were missed, the markup that follows would be counted as one endless section and refused.
        var tail = string.Concat(Enumerable.Repeat("<c><v>1</v></c>", 500));
        for (var filler = Edge - 30; filler < Edge + 30; filler++)
        {
            var lead = string.Concat(Enumerable.Repeat("<p/>", filler / 4)) + new string(' ', filler % 4);
            foreach (var close in new[] { "]]>", "]]]>", "]]]]]>" })
            {
                var text = "<a>" + lead + "<![CDATA[x" + close + tail + "</a>";
                Assert.Equal(text, ReadAll(Utf8(text), maxRun: limit));
                Assert.Equal(text, ReadAll(Utf8(text), maxRun: limit, chunk: 29));
            }
        }
    }

    [Fact]
    public void ManySmallCDataSections_EachStartAFreshCount()
    {
        var text = "<t>" + string.Concat(Enumerable.Repeat("<![CDATA[" + new string('<', 400) + "]]>", 300)) + "</t>";
        Assert.Equal(text, ReadAll(Utf8(text), maxRun: 500));
    }

    [Theory]
    [InlineData("]")]
    [InlineData("x]")]
    [InlineData("]]")]
    [InlineData("]]x")]
    [InlineData("]>")]
    public void CDataFullOfClosingBrackets_IsStillCounted(string unit)
    {
        // Content that keeps looking like the start of "]]>" without ever being it.
        var body = string.Concat(Enumerable.Repeat(unit, 3000 / unit.Length));
        var text = "<t><![CDATA[" + body + "x]]></t>";
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8(text), maxRun: 1000));
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8(text), maxRun: 1000, chunk: 7));
        Assert.Equal(text, ReadAll(Utf8(text), maxRun: 4000));
    }

    [Fact]
    public void CountAfterCData_StartsAtZero()
    {
        // Exactly the limit of text between the end of a section and the next tag is allowed, one more is not.
        const int limit = 500;
        var ok = "<t><![CDATA[x]]>" + new string('y', limit) + "</t>";
        var over = "<t><![CDATA[x]]>" + new string('y', limit + 1) + "</t>";
        Assert.Equal(ok, ReadAll(Utf8(ok), maxRun: limit));
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8(over), maxRun: limit));
        // And the section itself: exactly the limit of content is allowed.
        var full = "<t><![CDATA[" + new string('<', limit) + "]]></t>";
        var fuller = "<t><![CDATA[" + new string('<', limit + 1) + "]]></t>";
        Assert.Equal(full, ReadAll(Utf8(full), maxRun: limit));
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8(fuller), maxRun: limit));
    }

    [Fact]
    public void LessThanInsideCData_DoesNotResetTheCount()
    {
        // 400 + 400 + 400 with a '<' between: each stretch is under the limit, the section is not.
        var section = "<![CDATA[" + new string('a', 400) + "<" + new string('b', 400) + "<" + new string('c', 400) + "]]>";
        Assert.Throws<InvalidDataException>(() => ReadAll(Utf8("<t>" + section + "</t>"), maxRun: 1000));
    }

    [Fact]
    public void Utf16CharactersWhoseBytesLookLikeMarkup_AreCountedAsText()
    {
        // U+3C3C is two bytes of 0x3C, the byte of '<'. Counted on bytes, this attribute looked like markup.
        var text = "<c x=\"" + new string('㰼', 5000) + "\"/>";
        var bytes = Encoding.Unicode.GetPreamble().Concat(Encoding.Unicode.GetBytes(text)).ToArray();
        Assert.Throws<InvalidDataException>(() => ReadAll(bytes, maxRun: 1000));
        Assert.Equal(text, ReadAll(bytes, maxRun: 6000));
    }

    private static string Attributes(int count) => string.Concat(Enumerable.Range(0, count).Select(i => " a" + i + "=\"1\""));

    [Theory]
    [InlineData("<c")]                    // A plain tag of many attributes.
    [InlineData("<c z=\">\"")]            // The same behind a '>' in a quoted value, which does not end the tag.
    [InlineData("<c z='>'")]
    [InlineData("<c z='\"' y=\">\"")]     // A quote of the other kind inside a value does not open one.
    [InlineData("<c z=\"'\" y='>'")]
    public void TagOfManyAttributes_IsRefused_AndAShorterOneIsNot(string start)
    {
        // About 300,000 characters of attributes: nothing to a text limit of ten million, far over a tag limit of 70,000.
        var soup = Utf8("<r>" + start + Attributes(30_000) + "/></r>");
        var ex = Assert.Throws<InvalidDataException>(() => ReadAll(soup, maxRun: 10_000_000, maxTag: 70_000));
        Assert.Contains("tag longer than", ex.Message);
        Assert.Throws<InvalidDataException>(() => ReadAll(soup, maxRun: 10_000_000, maxTag: 70_000, chunk: 211));

        var modest = "<r>" + start + Attributes(3_000) + "/></r>";
        Assert.Equal(modest, ReadAll(Utf8(modest), maxRun: 10_000_000, maxTag: 70_000));
    }

    [Fact]
    public void TagLimit_DoesNotApplyToTextAfterAShortTag()
    {
        // Half a million characters of text full of quotes and '>' is text, not a tag.
        var body = string.Concat(Enumerable.Repeat("it's \"quoted\" > and so on ", 20_000));
        foreach (var open in new[] { "<t>", "<t xml:space=\"preserve\">", "<t a='x'>" })
        {
            var text = "<r>" + open + body + "</t></r>";
            Assert.Equal(text, ReadAll(Utf8(text), maxRun: 1_000_000, maxTag: 70_000));
            Assert.Equal(text, ReadAll(Utf8(text), maxRun: 1_000_000, maxTag: 70_000, chunk: 977));
        }
    }

    [Fact]
    public void TagLimit_IgnoresCommentsAndProcessingInstructions()
    {
        // Their content may hold quotes that never pair up; they are not tags and the parser skips them.
        var filler = new string('x', 200_000);
        var text = "<a><!-- it's a \"comment " + filler + " --><?pi it's " + filler + " ?><b/></a>";
        Assert.Equal(text, ReadAll(Utf8(text), maxRun: 1_000_000, maxTag: 70_000));
    }

    [Fact]
    public void TagLimit_AllowsOneLongAttributeValueUpToTheLimit()
    {
        var under = "<a><c sqref=\"" + new string('A', 40_000) + "\"/></a>";
        Assert.Equal(under, ReadAll(Utf8(under), maxRun: 1_000_000, maxTag: 70_000));
        var over = Utf8("<a><c sqref=\"" + new string('A', 120_000) + "\"/></a>");
        Assert.Throws<InvalidDataException>(() => ReadAll(over, maxRun: 1_000_000, maxTag: 70_000));
    }

    [Fact]
    public void TagLimit_FollowsATagThatStartsAtTheEdgeOfABuffer()
    {
        for (var padding = Edge - 12; padding < Edge + 12; padding++)
        {
            var lead = "<r>" + string.Concat(Enumerable.Repeat("<p/>", padding / 4)) + new string(' ', padding % 4);
            var soup = Utf8(lead + "<c" + Attributes(30_000) + "/></r>");
            Assert.Throws<InvalidDataException>(() => ReadAll(soup, maxRun: 10_000_000, maxTag: 70_000));
            Assert.Throws<InvalidDataException>(() => ReadAll(soup, maxRun: 10_000_000, maxTag: 70_000, chunk: 61));

            // A short tag in the same place, then long text: the tag ended, so only the text limit applies.
            var fine = lead + "<t a=\">\">" + new string('y', 200_000) + "</t></r>";
            Assert.Equal(fine, ReadAll(Utf8(fine), maxRun: 1_000_000, maxTag: 70_000));
        }
    }

    [Fact]
    public void XmlParser_ReadsThroughTheGuard_TheSameAsFromTheBytes()
    {
        var xml = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"urn:x\"><si><t xml:space=\"preserve\"> a &amp; b </t></si><si><t><![CDATA[c < d]]></t></si><!-- note --><si><r><t>é</t></r></si></sst>";
        foreach (var encoding in new Encoding[] { new UTF8Encoding(true), new UTF8Encoding(false), Encoding.Unicode, Encoding.BigEndianUnicode })
        {
            var declared = xml.Replace("UTF-8", encoding.WebName);
            var bytes = encoding.GetPreamble().Concat(encoding.GetBytes(declared)).ToArray();
            var settings = new XmlReaderSettings { IgnoreComments = true };
            Assert.Equal(Describe(XmlReader.Create(new MemoryStream(bytes), settings)), Describe(XmlReader.Create(new GuardedPartReader(new ChunkStream(bytes, 5), "p", long.MaxValue, true, 10_000, 10_000), settings)));
        }

        static string Describe(XmlReader reader)
        {
            var sb = new StringBuilder();
            using (reader)
                while (reader.Read())
                    if (reader.NodeType != XmlNodeType.XmlDeclaration) sb.Append(reader.NodeType).Append(':').Append(reader.Name).Append('=').Append(reader.Value).Append('|');
            return sb.ToString();
        }
    }

    [Fact]
    public void DepthGuard_StopsRead_Skip_AndSubtrees()
    {
        var deep = string.Concat(Enumerable.Repeat("<a>", 400)) + string.Concat(Enumerable.Repeat("</a>", 400));
        var xml = "<root><keep/>" + deep + "</root>";

        XmlReader Open()
        {
            var reader = new DepthGuardedXmlReader(XmlReader.Create(new StringReader(xml)));
            reader.Read();              // root
            reader.Read();              // keep
            reader.Read();              // the first nested <a>
            return reader;
        }

        using (var r = Open()) Assert.Throws<InvalidDataException>(() => { while (r.Read()) { } });
        using (var r = Open()) Assert.Throws<InvalidDataException>(() => r.Skip());
        using (var r = Open())
        using (var sub = r.ReadSubtree())
            Assert.Throws<InvalidDataException>(() => { while (sub.Read()) { } });

        // 250 levels is inside the limit and reads to the end.
        var fine = "<root>" + string.Concat(Enumerable.Repeat("<a>", 250)) + string.Concat(Enumerable.Repeat("</a>", 250)) + "</root>";
        using var ok = new DepthGuardedXmlReader(XmlReader.Create(new StringReader(fine)));
        var nodes = 0;
        while (ok.Read()) nodes++;
        Assert.Equal(502, nodes);
    }
}
