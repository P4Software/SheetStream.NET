using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// Regression tests for the 0.5.0 hardening work: limits on what the reader will unpack from an untrusted
/// file, values a damaged file can carry, Excel-exact rounding and date serials, in-place edits that only
/// commit on an explicit save, and the shared string memory budget.
/// </summary>
public class HardeningTests
{
    private const string SheetHead = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>";
    private const string SheetTail = "</sheetData></worksheet>";

    /// <summary>A one-sheet workbook whose sheet XML (or another part) is then replaced by hand-written markup.</summary>
    private static byte[] Seed() => ExcelPackage.CreateBytes(wb =>
    {
        var s = wb.Worksheets.Add("Data");
        s.Row().WriteCell("shared").WriteCell(TimeSpan.FromHours(1.5));
    });

    private static byte[] ReplacePart(byte[] package, string part, string content)
    {
        var ms = new MemoryStream();
        ms.Write(package, 0, package.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            zip.GetEntry(part)!.Delete();
            using var w = new StreamWriter(zip.CreateEntry(part).Open(), new UTF8Encoding(false));
            w.Write(content);
        }
        return ms.ToArray();
    }

    private static byte[] WithSheet(string rowsXml) => ReplacePart(Seed(), "xl/worksheets/sheet1.xml", SheetHead + rowsXml + SheetTail);

    private static List<ExcelRow> ReadAll(byte[] bytes, ReaderOptions? options = null)
    {
        using var reader = ExcelReader.Open(new MemoryStream(bytes), false, options);
        return reader.ReadRows(0).ToList();
    }

    // ------------------------------------------------------------------ reader limits

    [Fact]
    public void OversizedInlineCell_IsRejectedBeforeItIsMaterialised()
    {
        // 5 million characters in a file of a few kilobytes: the shape of a decompression bomb.
        var bytes = WithSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>" + new string('A', 5_000_000) + "</t></is></c></row>");
        Assert.True(bytes.Length < 100_000, "The test file should be small on the wire.");

        // Per-thread, not process-wide: other test classes run in parallel and allocate at the same time.
        var before = GC.GetAllocatedBytesForCurrentThread();
        var ex = Assert.Throws<InvalidDataException>(() => ReadAll(bytes));
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.Contains("MaxCellTextLength", ex.Message);
        // The default limit is 1,048,576 characters (about 2 MB); rejecting must not cost more than a few times that.
        Assert.True(allocated < 16 * 1024 * 1024, $"Rejecting the cell allocated {allocated:N0} bytes.");
    }

    [Fact]
    public void OversizedCell_IsReadWhenTheLimitIsRaised()
    {
        var bytes = WithSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>" + new string('A', 2_000_000) + "</t></is></c></row>");
        var rows = ReadAll(bytes, new ReaderOptions { MaxCellTextLength = 3_000_000 });
        Assert.Equal(2_000_000, ((string)rows[0][0]!).Length);
    }

    [Fact]
    public void CellLimit_AppliesToFormulaStringsAndRichTextRuns()
    {
        var options = new ReaderOptions { MaxCellTextLength = 100 };
        var longText = new string('x', 80);
        // Two runs of 80 are over the limit together although each is under it.
        var rich = WithSheet($"<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><r><t>{longText}</t></r><r><t>{longText}</t></r></is></c></row>");
        Assert.Throws<InvalidDataException>(() => ReadAll(rich, options));
        var str = WithSheet($"<row r=\"1\"><c r=\"A1\" t=\"str\"><f>A2</f><v>{longText}{longText}</v></c></row>");
        Assert.Throws<InvalidDataException>(() => ReadAll(str, options));
        // Within the limit both read normally.
        var ok = WithSheet($"<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><r><t>{longText}</t></r></is></c><c r=\"B1\" t=\"str\"><v>{longText}</v></c></row>");
        var row = ReadAll(ok, options)[0];
        Assert.Equal(longText, row[0]);
        Assert.Equal(longText, row[1]);
    }

    [Fact]
    public void RowTextLimit_BoundsARowOfManyCellsEachUnderTheCellLimit()
    {
        var cell = new string('y', 400);
        var sb = new StringBuilder("<row r=\"1\">");
        for (var c = 0; c < 50; c++) sb.Append("<c t=\"inlineStr\"><is><t>").Append(cell).Append("</t></is></c>");
        sb.Append("</row>");
        var bytes = WithSheet(sb.ToString());

        var ex = Assert.Throws<InvalidDataException>(() => ReadAll(bytes, new ReaderOptions { MaxCellTextLength = 1000, MaxRowTextLength = 10_000 }));
        Assert.Contains("MaxRowTextLength", ex.Message);
        Assert.Equal(50, ReadAll(bytes, new ReaderOptions { MaxCellTextLength = 1000, MaxRowTextLength = 30_000 })[0].Values.Count);
    }

    [Fact]
    public void RowTextLimit_DoesNotCountSharedStrings()
    {
        // 2,000 cells all pointing at one 600-character shared string: 1.2 million characters on screen, one string in memory.
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var row = wb.Worksheets.Add("Data").Row();
            var text = new string('z', 600);
            for (var c = 0; c < 2000; c++) row.WriteCell(text);
        });
        var rows = ReadAll(bytes, new ReaderOptions { MaxCellTextLength = 1000, MaxRowTextLength = 5000 });
        Assert.Equal(2000, rows[0].Values.Count);
    }

    [Fact]
    public void SharedStringTable_IsCappedByLoadedPartBytes()
    {
        var sb = new StringBuilder("<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        for (var i = 0; i < 20_000; i++) sb.Append("<si><t>string number ").Append(i).Append("</t></si>");
        sb.Append("</sst>");
        var bytes = ReplacePart(Seed(), "xl/sharedStrings.xml", sb.ToString());

        var ex = Assert.Throws<InvalidDataException>(() => ReadAll(bytes, new ReaderOptions { MaxLoadedPartBytes = 100_000 }));
        Assert.Contains("MaxLoadedPartBytes", ex.Message);
        Assert.NotEmpty(ReadAll(bytes)); // The default limit is far above this table.
    }

    [Fact]
    public void OversizedSharedString_IsRejected()
    {
        var sst = "<?xml version=\"1.0\" encoding=\"UTF-8\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><si><t>" + new string('B', 5000) + "</t></si></sst>";
        var bytes = ReplacePart(Seed(), "xl/sharedStrings.xml", sst);
        Assert.Throws<InvalidDataException>(() => ReadAll(bytes, new ReaderOptions { MaxCellTextLength = 1000 }));
    }

    [Fact]
    public void OversizedAttribute_IsRejected()
    {
        // The text limits cannot see an attribute; the token guard does.
        var bytes = WithSheet("<row r=\"1\"><c r=\"A1\" junk=\"" + new string('Q', 400_000) + "\"><v>1</v></c></row>");
        var ex = Assert.Throws<InvalidDataException>(() => ReadAll(bytes, new ReaderOptions { MaxCellTextLength = 1000 }));
        Assert.Contains("longer than", ex.Message);
    }

    [Fact]
    public void AbsurdNesting_IsRejected_InRowsAndWhenSkippingThem()
    {
        var deep = new StringBuilder("<row r=\"1\"><c r=\"A1\"><f>");
        for (var i = 0; i < 5000; i++) deep.Append("<a>");
        for (var i = 0; i < 5000; i++) deep.Append("</a>");
        deep.Append("</f><v>1</v></c></row>");
        var bytes = WithSheet(deep.ToString());

        Assert.Throws<InvalidDataException>(() => ReadAll(bytes));
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        Assert.Throws<InvalidDataException>(() => reader.ReadMergedRanges(0)); // Skips sheetData, still depth-checked.
    }

    [Fact]
    public async Task AsyncReader_EnforcesTheSameLimits()
    {
        var bytes = WithSheet("<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>" + new string('A', 5000) + "</t></is></c></row>");
        var options = new ReaderOptions { MaxCellTextLength = 1000 };
        using var reader = await ExcelReader.OpenAsync(new MemoryStream(bytes), false, options);
        await Assert.ThrowsAsync<InvalidDataException>(async () =>
        {
            await foreach (var _ in reader.ReadRowsAsync(0)) { }
        });

        using var lenient = await ExcelReader.OpenAsync(new MemoryStream(bytes), false, new ReaderOptions { MaxCellTextLength = 10_000 });
        var rows = new List<ExcelRow>();
        await foreach (var row in lenient.ReadRowsAsync(0)) rows.Add(row);
        Assert.Equal(5000, ((string)rows[0][0]!).Length);
    }

    [Fact]
    public async Task ChunkedTextReading_ReturnsTheSameTextAsBefore()
    {
        // Values around the 1,024-character chunk size, with entities, whitespace to preserve and escapes.
        string[] texts = { "", "a", " leading and trailing ", "a & b < c > d \"q\"", new string('k', 1023), new string('k', 1024), new string('k', 1025), new string('é', 5000) + "\ttab\nnewline", "ctl\u0001char _x0041_ literal" };
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            foreach (var t in texts) s.Row().WriteCell(t).WriteInlineString(t).WriteRichText(new RichText(t) { Bold = true }, "|", t);
        });
        var rows = ReadAll(bytes);
        using var asyncReader = await ExcelReader.OpenAsync(new MemoryStream(bytes));
        var asyncRows = new List<ExcelRow>();
        await foreach (var row in asyncReader.ReadRowsAsync(0)) asyncRows.Add(row);

        for (var i = 0; i < texts.Length; i++)
        {
            foreach (var row in new[] { rows[i], asyncRows[i] })
            {
                Assert.Equal(texts[i].Length == 0 ? null : texts[i], row[0] is string s0 && s0.Length == 0 ? null : row[0]);
                Assert.Equal(texts[i], row[1] as string ?? "");
                Assert.Equal(texts[i] + "|" + texts[i], row[2]);
            }
        }
    }

    [Fact]
    public void NonSeekableStream_SpillsToATempFileAboveTheMemoryLimit()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            for (var i = 0; i < 2000; i++) s.Row().WriteCell("Row " + i).WriteCell(i);
        });
        using (var spilled = ExcelReader.Buffer(new ForwardOnly(bytes), memoryLimit: 4096))
        {
            var file = Assert.IsType<FileStream>(spilled);
            Assert.Equal(bytes.Length, file.Length);
            Assert.Equal(0, file.Position);
            var path = file.Name;
            Assert.True(File.Exists(path));
            using (var reader = ExcelReader.Open(spilled)) Assert.Equal(2000, reader.ReadRows(0).Count());
            Assert.False(File.Exists(path)); // Deleted with the reader.
        }
        using var small = ExcelReader.Buffer(new ForwardOnly(bytes), memoryLimit: long.MaxValue);
        Assert.IsType<MemoryStream>(small);

        using var viaOpen = ExcelReader.Open(new ForwardOnly(bytes));
        Assert.Equal(2000, viaOpen.ReadRows(0).Count());
    }

    [Fact]
    public async Task ExistingCallsWithDefaultLiterals_StillCompileAndBindToTheOldOverloads()
    {
        // Source compatibility with 0.4.0: a bare `default` argument must not become ambiguous now that
        // overloads taking ReaderOptions exist. Compiling this test is the assertion.
        var bytes = Seed();
        var path = Path.Combine(Path.GetTempPath(), "ss-compat-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.WriteAllBytes(path, bytes);
        try
        {
            using (var a = await ExcelReader.OpenAsync(path, default)) Assert.Single(a.SheetNames);
            using (var b = await ExcelReader.OpenAsync(new MemoryStream(bytes), false, default)) Assert.Single(b.SheetNames);
            using (var c = await ExcelReader.OpenAsync(new MemoryStream(bytes), cancellationToken: default)) Assert.Single(c.SheetNames);
            using (var d = ExcelReader.Open(new MemoryStream(bytes), false)) Assert.Single(d.SheetNames);
            using (var e = ExcelReader.Open(path, null)) Assert.Single(e.SheetNames);
        }
        finally
        {
            File.Delete(path);
        }
    }

    // ------------------------------------------------------------------ values a damaged file can carry

    [Theory]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("-Infinity")]
    [InlineData("1E300")]
    [InlineData("-1E300")]
    [InlineData("1E400")]
    public void HostileNumberInADurationOrDateCell_ReadsWithoutThrowing(string value)
    {
        // Style 1 of the seed is the duration format of its TimeSpan cell.
        var bytes = WithSheet($"<row r=\"1\"><c r=\"A1\" s=\"1\"><v>{value}</v></c><c r=\"B1\"><v>{value}</v></c></row>");
        var row = ReadAll(bytes)[0];
        Assert.NotNull(row[0]);
        Assert.IsNotType<TimeSpan>(row[0]);
        _ = row.GetString(0);
    }

    [Fact]
    public void OrdinaryDuration_StillReadsAsTimeSpan()
    {
        var row = ReadAll(Seed())[0];
        Assert.Equal(TimeSpan.FromHours(1.5), row[1]);
    }

    // ------------------------------------------------------------------ Excel-exact results

    [Theory]
    // Expected values are what Excel 16 returns for =ROUND(value, digits).
    [InlineData(1.005, 2, 1.01)]
    [InlineData(0.285, 2, 0.29)]
    [InlineData(2.675, 2, 2.68)]
    [InlineData(-1.005, 2, -1.01)]
    [InlineData(2.5, 0, 3)]
    [InlineData(-2.5, 0, -3)]
    [InlineData(1234.5678, -2, 1200)]
    [InlineData(1250, -2, 1300)]
    [InlineData(0.1 + 0.2, 15, 0.3)]
    public void CachedRound_MatchesExcel(double value, int digits, double expected)
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.CacheFormulaResults();
            s.Row().WriteCell(value).WriteFormula($"ROUND(A1,{digits})");
        });
        Assert.Equal(expected, ReadAll(bytes)[0][1]);
    }

    [Fact]
    public void CachedRound_OfANumberTooLargeToRoundExactly_WritesNoValue()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.CacheFormulaResults();
            s.Row().WriteCell(1e40).WriteFormula("ROUND(A1,-3)");
        });
        Assert.Null(ReadAll(bytes)[0][1]); // No cached value rather than a wrong one.
    }

    [Theory]
    // Excel's serials: 1 = 1 January 1900, 59 = 28 February, 61 = 1 March (60 is its fictitious 29 February).
    [InlineData(1900, 1, 1, 1)]
    [InlineData(1900, 2, 28, 59)]
    [InlineData(1900, 3, 1, 61)]
    [InlineData(2026, 11, 13, 46339)]
    public void Dates_AreWrittenWithExcelsSerial_AndReadBack(int year, int month, int day, double serial)
    {
        var date = new DateTime(year, month, day);
        var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Data").Row().WriteCell(date));

        using (var zip = new ZipArchive(new MemoryStream(bytes)))
        using (var sr = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open()))
            Assert.Contains($"<v>{serial}</v>", sr.ReadToEnd());

        Assert.Equal(date, ReadAll(bytes)[0][0]);
    }

    // ------------------------------------------------------------------ in-place edit

    [Fact]
    public void EditInPlace_WithoutSave_LeavesTheOriginalUntouched()
    {
        var dir = Path.Combine(Path.GetTempPath(), "ss-hard-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "book.xlsx");
        try
        {
            var original = Seed();
            File.WriteAllBytes(path, original);

            // The caller's code fails between two changes that belong together.
            // A lambda that only throws also converts to Func<Task>; naming the delegate picks the synchronous overload.
            var ex = Record.Exception((Action)(() =>
            {
                using var package = ExcelPackage.Edit(path);
                package.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "half an edit";
                throw new TimeoutException("the second change never happened");
            }));
            Assert.IsType<TimeoutException>(ex); // And the caller sees its own exception, not one from a save.
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(dir)); // No temporary file left behind.

            // Simply forgetting Save behaves the same way.
            using (var package = ExcelPackage.Edit(path))
                package.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "forgot to save";
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(dir));

            using (var package = ExcelPackage.Edit(path))
            {
                package.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "saved";
                package.Save();
            }
            using var reader = ExcelReader.Open(path);
            Assert.Equal("saved", reader.ReadRows(0).First()[0]);
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    // ------------------------------------------------------------------ shared string budget

    [Fact]
    public void SharedStringBudget_StopsTheTableGrowing_AndKeepsTheFileValid()
    {
        var ms = new MemoryStream();
        int unique;
        using (var package = new ExcelPackage(ms, leaveOpen: true, new PackageOptions { MaxSharedStringChars = 5000 }))
        {
            var s = package.Workbook.Worksheets.Add("Data");
            for (var i = 0; i < 2000; i++) s.Row().WriteCell("Unique value " + i).WriteCell("Repeated");
            package.Save();
            unique = package.Workbook.SharedStrings.UniqueCount;
        }

        Assert.InRange(unique, 2, 200); // 5,000 units at about 45 per entry: roughly a hundred strings, not 2,001.
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, false))
            Assert.Empty(new OpenXmlValidator().Validate(doc));

        var rows = ReadAll(ms.ToArray());
        Assert.Equal(2000, rows.Count);
        for (var i = 0; i < rows.Count; i++)
        {
            Assert.Equal("Unique value " + i, rows[i][0]);
            Assert.Equal("Repeated", rows[i][1]); // In the table from the first row, so shared throughout.
        }
    }

    [Fact]
    public void SharedStringBudget_DefaultIsNotReachedByOrdinaryExports()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            for (var i = 0; i < 5000; i++) s.Row().WriteCell("Customer " + i);
            Assert.Equal(5000, wb.SharedStrings.UniqueCount);
        });
        Assert.Equal(5000, ReadAll(bytes).Count);
    }

    /// <summary>A stream that cannot seek, like an HTTP request body.</summary>
    private sealed class ForwardOnly : Stream
    {
        private readonly MemoryStream _inner;
        public ForwardOnly(byte[] bytes) => _inner = new MemoryStream(bytes);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
