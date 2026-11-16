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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AFailedWriteToTheOutput_NeverEndsInASavedFile(bool autoFit)
    {
        var output = new FlakyStream(failOnWrite: 3);
        var package = new ExcelPackage(output, leaveOpen: true, new PackageOptions { CompressionLevel = CompressionLevel.Fastest });
        var sheet = package.Workbook.Worksheets.Add("S");
        if (autoFit) sheet.AutoFitColumns();

        // A caller that catches the failure and carries on, which is what a retry loop or a lenient export does.
        var failures = 0;
        for (var i = 0; i < 400_000; i++)
        {
            try { sheet.Row().WriteCell(i).WriteCell("row " + (i * 7919L)).WriteCell(i * 0.37); }
            catch (Exception) { failures++; }
        }

        if (autoFit)
        {
            // The rows only go to the zip entry at the end, so that is where the failure is met.
            Assert.ThrowsAny<Exception>(() => package.Save());
            return;
        }
        Assert.True(failures > 0, "The write failure never reached the caller.");
        Assert.ThrowsAny<Exception>(() => package.Save());
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
