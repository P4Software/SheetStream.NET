using System.IO.Compression;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// What happens when something goes wrong while a sheet is being written. Rows used to go through the part's
/// <c>XmlWriter</c>, which is dead after any failure (every later call throws), so a failed export could never be
/// saved as a file with a hole in it. Rows are now composed in the zip entry's buffer, and these tests pin that the
/// promise survived: a failure is never silently turned into a damaged file.
/// </summary>
public class FailStopTests
{
    /// <summary>A write-only stream that throws once, on the n-th write, and then behaves.</summary>
    private sealed class FlakyStream : Stream
    {
        private readonly int _failOnWrite;
        private int _writes;

        public FlakyStream(int failOnWrite) => _failOnWrite = failOnWrite;

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
            if (++_writes == _failOnWrite) throw new IOException("The disk is full.");
        }
    }

    /// <summary>A write-only stream that works until it has taken a given number of bytes and fails on every write after that.</summary>
    private sealed class FailsForGoodStream : Stream
    {
        private readonly long _allowed;
        private long _written;

        public FailsForGoodStream(long allowedBytes) => _allowed = allowedBytes;

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
    /// Writes <paramref name="rows"/> rows the way a caller that swallows errors does, and returns every exception
    /// the row calls threw, in order.
    /// </summary>
    private static List<Exception> WriteRowsSwallowingErrors(ExcelPackage package, bool autoFit, int rows)
    {
        var sheet = package.Workbook.Worksheets.Add("S");
        if (autoFit) sheet.AutoFitColumns();
        var thrown = new List<Exception>();
        for (var i = 0; i < rows; i++)
        {
            try { sheet.Row().WriteCell(i).WriteCell("row " + (i * 7919L)).WriteCell(i * 0.37); }
            catch (Exception ex) { thrown.Add(ex); }
        }
        return thrown;
    }

    /// <summary>
    /// The failure the caller meets is the failure of the output: an <see cref="IOException"/>, the first time as the
    /// stream threw it and every time after as the library's report that the part cannot be completed. Never an
    /// index or null-reference error from the writer's own bookkeeping, which is what a half-reset buffer used to produce.
    /// </summary>
    private static void AssertOnlyIOExceptions(List<Exception> thrown)
    {
        Assert.NotEmpty(thrown);
        // A summary instead of Assert.All: a broken writer throws on every one of tens of thousands of rows, and a
        // failure message that lists each of them is too big for the test runner to report.
        var others = thrown.Where(ex => ex.GetType() != typeof(IOException)).ToList();
        Assert.True(others.Count == 0,
            $"{others.Count} of {thrown.Count} exceptions were not IOException. First: {others.FirstOrDefault()?.GetType().Name}: {others.FirstOrDefault()?.Message}");
    }

    /// <summary>Save fails with an I/O error that still says what the output reported (here: the full disk), not with a bookkeeping error.</summary>
    private static void AssertSaveReportsTheOriginalFailure(ExcelPackage package)
    {
        var failure = Assert.IsType<IOException>(Assert.ThrowsAny<Exception>(() => package.Save()));
        Assert.Contains("The disk is full.", failure.Message);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AFailedWriteToTheOutput_NeverEndsInASavedFile(bool autoFit)
    {
        var output = new FlakyStream(failOnWrite: 3);
        var package = new ExcelPackage(output, leaveOpen: true, new PackageOptions { CompressionLevel = CompressionLevel.Fastest });

        // A caller that catches the failure and carries on, which is what a retry loop or a lenient export does.
        var thrown = WriteRowsSwallowingErrors(package, autoFit, 100_000);

        if (autoFit)
        {
            // The rows only go to the zip entry at the end, so that is where the failure is met.
            Assert.IsType<IOException>(Assert.ThrowsAny<Exception>(() => package.Save()));
            return;
        }
        Assert.True(thrown.Count > 0, "The write failure never reached the caller.");
        AssertOnlyIOExceptions(thrown);
        AssertSaveReportsTheOriginalFailure(package);
    }

    /// <summary>
    /// An output that dies for good part-way through. The first failed row call reports the stream's own error; every
    /// row after it, 20,000 of them each in its own try block, must report the same kind of failure, and so must Save.
    /// </summary>
    [Theory]
    [InlineData(false, 70_000)]
    [InlineData(false, 70_001)]
    [InlineData(false, 123_457)]
    [InlineData(true, 70_000)]
    public void AnOutputThatFailsForGood_ReportsAnIOExceptionOnEveryLaterRowAndOnSave(bool autoFit, long allowedBytes)
    {
        var package = new ExcelPackage(new FailsForGoodStream(allowedBytes), leaveOpen: true, new PackageOptions { CompressionLevel = CompressionLevel.NoCompression });

        var thrown = WriteRowsSwallowingErrors(package, autoFit, 20_000);

        // With AutoFit the rows go to a temp file and the output is met when the sheet is sealed, at Save.
        if (!autoFit)
        {
            AssertOnlyIOExceptions(thrown);
            Assert.True(thrown.Count > 1, "Only the first row call reported the failure.");
        }
        else
        {
            Assert.Empty(thrown);
        }
        if (autoFit) Assert.IsType<IOException>(Assert.ThrowsAny<Exception>(() => package.Save()));
        else AssertSaveReportsTheOriginalFailure(package);
        Assert.False(package.IsSaved);
    }

    /// <summary>
    /// A rich-text run whose font name XML cannot carry fails while it is rendered. Nothing of the cell may be
    /// left in the sheet: the file stays well formed, the cell is simply not there.
    /// </summary>
    [Fact]
    public void ARichTextCellThatFailsToRender_LeavesNothingBehind()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteCell("before");
            Assert.ThrowsAny<ArgumentException>(() => s.Row().WriteRichText(new RichText("bad") { FontName = "a\u0001b" }));
            s.Row().WriteCell("after");
        });

        // Parsed by a real XML reader: a half-written cell would make the sheet not well formed.
        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using (var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()) System.Xml.Linq.XDocument.Load(stream);

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("S").ToList();
        Assert.Equal("before", rows[0][0]);
        Assert.Equal("after", rows[^1][0]);
        Assert.DoesNotContain(rows, r => r[0] is string s && s == "bad");
    }
}
