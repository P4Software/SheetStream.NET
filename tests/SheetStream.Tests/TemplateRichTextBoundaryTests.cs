using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// Rich text on sheets that are filled from a template or edited. Those sheets hand their rows to an XmlWriter as
/// text, so the row writer decodes its UTF-8 staging buffer back to characters in 16 KB blocks. A rich-text cell's
/// bytes are copied into that buffer in pieces, so a block can end in the middle of a character, and the decoder
/// then carries the unfinished bytes over to the next block. These tests put characters of every UTF-8 length on
/// each side of that block boundary.
/// </summary>
public class TemplateRichTextBoundaryTests
{
    /// <summary>One character of each UTF-8 length above ASCII: 2 bytes, 3 bytes, and 4 bytes (a surrogate pair).</summary>
    public static IEnumerable<object[]> WideCharacters() => new[]
    {
        new object[] { "é" },
        new object[] { "€" },
        new object[] { "\U0001F600" },
    };

    private static byte[] EmptyTemplate() => ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S"));

    private static byte[] OneCellWorkbook() => ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S").Row(1).WriteCell("x"));

    /// <summary>Parses every worksheet part with a real XML parser and returns the text of the one cell it should hold.</summary>
    private static void AssertWellFormedAndReadsBack(byte[] xlsx, string expected)
    {
        using (var zip = new ZipArchive(new MemoryStream(xlsx), ZipArchiveMode.Read))
        {
            foreach (var entry in zip.Entries.Where(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.Ordinal) && e.FullName.EndsWith(".xml", StringComparison.Ordinal)))
                using (var stream = entry.Open()) XDocument.Load(stream);
        }

        using var reader = ExcelReader.Open(new MemoryStream(xlsx));
        var texts = reader.ReadRows("S").SelectMany(r => r.Values).OfType<string>().Where(s => s != "x" && s != Filler).ToList();
        var actual = Assert.Single(texts);
        // Compared without Assert.Equal on the strings themselves: a failure would print 36,000 characters.
        Assert.True(actual.Length == expected.Length && string.CompareOrdinal(actual, expected) == 0,
            $"The text read back differs from the text written (length {actual.Length} against {expected.Length}).");
    }

    /// <summary>
    /// Text with the wide character at <paramref name="offset"/> and 16,000 ASCII characters after it. A cell may hold at most
    /// 32,767 characters, so the text alone cannot make the block after the first flush a full 16 KB (the rest of it is about
    /// 16,030 bytes); <see cref="Filler"/> in the next cell of the row does. A full block is what makes the decoder, which
    /// carries up to three bytes over from the flush, produce one character more than the buffer has bytes (the bug these
    /// tests guard).
    /// </summary>
    private static string TextWithWideCharacterAt(int offset, string wide) => new string('a', offset) + wide + new string('b', 16_000);

    /// <summary>The cell written after the rich text, so that the second block handed to the decoder is a whole buffer.</summary>
    private static readonly string Filler = new string('c', 2_000);

    /// <summary>Offsets around the first flush: the cell's 27 bytes of markup plus the text fill the 16 KB buffer near 16,350.</summary>
    private static IEnumerable<int> Offsets() => Enumerable.Range(16_330, 71);

    [Theory]
    [MemberData(nameof(WideCharacters))]
    public void TemplateSheet_RichTextWithAWideCharacterAtEveryOffsetAroundTheFlush_IsWrittenWhole(string wide)
    {
        var template = EmptyTemplate();
        foreach (var offset in Offsets())
        {
            var text = TextWithWideCharacterAt(offset, wide);
            using var output = new MemoryStream();
            using (var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true))
            {
                package.Workbook.Worksheets.Open("S").Row(1).WriteRichText(new RichText(text)).WriteInlineString(Filler);
                package.Save();
            }
            AssertWellFormedAndReadsBack(output.ToArray(), text);
        }
    }

    [Theory]
    [MemberData(nameof(WideCharacters))]
    public void EditedSheet_RichTextWithAWideCharacterAtEveryOffsetAroundTheFlush_IsWrittenWhole(string wide)
    {
        var input = OneCellWorkbook();
        foreach (var offset in Offsets())
        {
            var text = TextWithWideCharacterAt(offset, wide);
            using var output = new MemoryStream();
            using (var package = ExcelPackage.Edit(new MemoryStream(input), output, leaveOpen: true))
            {
                {
                var edited = package.Workbook.Worksheets.Edit("S");
                edited.Cells[2, 1].Value = new[] { new RichText(text) };
                edited.Cells[2, 2].Value = Filler;
            }
                package.Save();
            }
            AssertWellFormedAndReadsBack(output.ToArray(), text);
        }
    }

    /// <summary>A write-only stream that works until it has taken a given number of bytes and fails on every write after that.</summary>
    private sealed class FailsAfterStream : Stream
    {
        private readonly long _allowed;
        private long _written;

        public FailsAfterStream(long allowedBytes) => _allowed = allowedBytes;

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_written + count > _allowed) throw new IOException("The disk is full.");
            _written += count;
        }
    }

    /// <summary>
    /// The output fails in the middle of a rich-text cell. A caller that catches the error, writes more and saves must
    /// never be told the save worked: either Save throws, or what it wrote is a well-formed sheet.
    /// </summary>
    [Fact]
    public void TemplateSheet_FailureInTheMiddleOfACell_NeverEndsInASilentlyBrokenSave()
    {
        var text = TextWithWideCharacterAt(16_354, "\U0001F600");
        var options = new PackageOptions { CompressionLevel = CompressionLevel.NoCompression };
        using var package = ExcelPackage.OpenTemplate(new MemoryStream(EmptyTemplate()), new FailsAfterStream(allowedBytes: 9_000), leaveOpen: true, options);
        var sheet = package.Workbook.Worksheets.Open("S");

        // A caller that catches the failure and carries on.
        var failures = 0;
        for (var row = 1; row <= 6; row++)
        {
            try { sheet.Row(row).WriteRichText(new RichText(text)); }
            catch (Exception) { failures++; }
        }
        Assert.True(failures > 0, "The write failure never reached the caller.");

        // The output keeps failing, so a save cannot have worked; what matters is that it says so.
        Assert.ThrowsAny<Exception>(() => package.Save());
        Assert.False(package.IsSaved);
    }

    /// <summary>
    /// A cell that fails while it is being written, on a stream that works: the caller catches the error and
    /// carries on. Whatever Save does next, it may not report success for a sheet with half a cell in it.
    /// </summary>
    [Theory]
    [MemberData(nameof(WideCharacters))]
    public void TemplateSheet_ACaughtFailureInOneCell_NeverLeavesAMalformedSheetBehindASuccessfulSave(string wide)
    {
        var template = EmptyTemplate();
        foreach (var offset in Offsets())
        {
            var text = TextWithWideCharacterAt(offset, wide);
            using var output = new MemoryStream();
            using var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true);
            var sheet = package.Workbook.Worksheets.Open("S");
            try { sheet.Row(1).WriteRichText(new RichText(text)).WriteInlineString(Filler); }
            catch (Exception) { /* The caller swallows it, on purpose. */ }
            sheet.Row(2).WriteCell("after");

            bool saved;
            try { package.Save(); saved = true; }
            catch (Exception) { saved = false; }

            if (!saved) continue; // Refusing to save is a fine answer.
            using var zip = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
            using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XDocument.Load(stream); // Throws when the sheet is malformed.
        }
    }
}
