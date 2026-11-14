using System.Diagnostics;
using System.IO.Compression;
using System.Text;
using SheetStream.Internal;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// One test for each problem a separate adversarial review pass found in the 0.6.0 hardening work before it was released.
/// The review attacked the library through its public API with files built for the purpose; each test here
/// rebuilds the file that broke it. They are kept together because they share a lesson: bounding memory was
/// not the whole job. Time (a path resolved in quadratic time, a format code scanned once per cell), the zip
/// container itself, and the things callers collect all needed limits of their own.
/// </summary>
public class ReviewFindingsTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";

    private static byte[] Seed() => ExcelPackage.CreateBytes(wb =>
    {
        var s = wb.Worksheets.Add("Data");
        s.Row().WriteCell("Name").WriteCell("Qty");
        s.Row().WriteCell("a").WriteCell(1);
    });

    /// <summary>The seed workbook with some parts replaced, and optionally some added.</summary>
    private static byte[] With(params (string Part, string Content)[] parts)
    {
        var output = new MemoryStream();
        using (var from = new ZipArchive(new MemoryStream(Seed()), ZipArchiveMode.Read))
        using (var to = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            var done = new HashSet<string>();
            foreach (var entry in from.Entries)
            {
                using var target = to.CreateEntry(entry.FullName, CompressionLevel.Fastest).Open();
                var replacement = parts.FirstOrDefault(p => p.Part == entry.FullName);
                if (replacement.Part == null)
                {
                    using var source = entry.Open();
                    source.CopyTo(target);
                    continue;
                }
                done.Add(entry.FullName);
                using var writer = new StreamWriter(target, new UTF8Encoding(false));
                writer.Write(replacement.Content);
            }
            foreach (var (part, content) in parts.Where(p => !done.Contains(p.Part)))
            {
                using var writer = new StreamWriter(to.CreateEntry(part, CompressionLevel.Fastest).Open(), new UTF8Encoding(false));
                writer.Write(content);
            }
        }
        return output.ToArray();
    }

    private static string Sheet(string rows) => $"<worksheet xmlns=\"{Main}\"><sheetData>{rows}</sheetData></worksheet>";

    /// <summary>Runs <paramref name="action"/> and fails if it takes longer than <paramref name="seconds"/>.</summary>
    private static Exception? Timed(double seconds, Action action, string what)
    {
        var watch = Stopwatch.StartNew();
        var ex = Record.Exception(action);
        Assert.True(watch.Elapsed.TotalSeconds < seconds, $"{what} took {watch.Elapsed.TotalSeconds:F1} s; it must not take time out of proportion to the file.");
        return ex;
    }

    private static void ReadRows(byte[] bytes)
    {
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        foreach (var name in reader.SheetNames)
            foreach (var row in reader.ReadRows(name)) _ = row.Values.Count;
    }

    private static void EditAndSave(byte[] bytes)
    {
        using var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true);
        package.Workbook.Worksheets.Edit("Data").Cells["C5"].Value = 1;
        package.Save();
    }

    private static void CopyAndSave(byte[] bytes, PackageOptions? options = null)
    {
        using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true, options);
        package.Save();
    }

    // ------------------------------------------------------------------ time

    [Fact]
    public void RelationshipTargetOfManyDotSegments_ResolvesInLinearTime()
    {
        // Half a million "./" segments are six kilobytes once zipped. Removing them from the front of a list
        // one at a time took 43 seconds.
        var target = string.Concat(Enumerable.Repeat("./", 400_000)) + "worksheets/sheet1.xml";
        var rels = $"<Relationships xmlns=\"{PackageRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/worksheet\" Target=\"{target}\"/><Relationship Id=\"rId2\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"{Rel}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>";
        var bytes = With(("xl/_rels/workbook.xml.rels", rels));
        Assert.True(bytes.Length < 20_000);

        Assert.Null(Timed(5, () => ReadRows(bytes), "Reading"));
        Assert.Null(Timed(5, () => EditAndSave(bytes), "Editing"));
    }

    [Theory]
    [InlineData("", "a/b.xml", "a/b.xml")]
    [InlineData("xl/", "worksheets/sheet1.xml", "xl/worksheets/sheet1.xml")]
    [InlineData("xl/worksheets/", "../media/image1.png", "xl/media/image1.png")]
    [InlineData("xl/", "/docProps/app.xml", "docProps/app.xml")]
    [InlineData("xl/", "./././styles.xml", "xl/styles.xml")]
    [InlineData("xl/", "../../../../etc/passwd", "etc/passwd")]      // Never climbs out of the package.
    [InlineData("xl/", "a//b/./c/../d", "xl/a/b/d")]
    [InlineData("xl/", "/", "")]
    [InlineData("xl/", "..", "")]
    public void Resolve_NormalisesATarget(string folder, string target, string expected)
        => Assert.Equal(expected, TemplateSource.Resolve(folder, target));

    [Fact]
    public void VeryLongNumberFormat_IsInspectedOncePerFormat_NotOncePerCellFormatOrCell()
    {
        // One format code of half a million characters, used by thousands of cell formats and by every cell.
        var code = new string('0', 500_000);
        var xfs = string.Concat(Enumerable.Repeat("<xf numFmtId=\"164\"/>", 4000));
        var styles = $"<styleSheet xmlns=\"{Main}\"><numFmts count=\"1\"><numFmt numFmtId=\"164\" formatCode=\"{code}\"/></numFmts><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"4000\">{xfs}</cellXfs></styleSheet>";
        var rows = string.Concat(Enumerable.Range(1, 3000).Select(r => $"<row r=\"{r}\"><c r=\"A{r}\" s=\"1\"><v>{r}</v></c></row>"));
        var bytes = With(("xl/styles.xml", styles), ("xl/worksheets/sheet1.xml", Sheet(rows)));
        Assert.True(bytes.Length < 60_000);

        Assert.Null(Timed(8, () => ReadRows(bytes), "Reading"));
        Assert.Null(Timed(8, () => EditAndSave(bytes), "Editing"));
    }

    [Fact]
    public void DataTableHeaderOfIdenticalNames_IsNamedInLinearTime()
    {
        var header = "<row r=\"1\">" + string.Concat(Enumerable.Repeat("<c t=\"inlineStr\"><is><t>a</t></is></c>", 16_384)) + "</row>";
        var bytes = With(("xl/worksheets/sheet1.xml", Sheet(header + "<row r=\"2\"><c r=\"A2\"><v>1</v></c></row>")));
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        System.Data.DataTable? table = null;
        Assert.Null(Timed(8, () => table = reader.ReadDataTable("Data"), "Building the table"));
        Assert.Equal(16_384, table!.Columns.Count);
        Assert.Equal("a", table.Columns[0].ColumnName);
        Assert.Equal("a2", table.Columns[1].ColumnName);
        Assert.Equal("a16384", table.Columns[16_383].ColumnName);
    }

    [Fact]
    public void RangeListOfTokensThatAreNotRanges_IsStretchedWithoutAnExceptionPerToken()
    {
        // Filling a template stretches its conditional formats over the new rows, which means reading every
        // token of their range lists. 450,000 tokens that are not ranges used to cost an exception each.
        var junk = string.Join(" ", Enumerable.Repeat("x", 450_000));
        var sheet = $"<worksheet xmlns=\"{Main}\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>"
            + string.Concat(Enumerable.Range(0, 4).Select(_ => $"<conditionalFormatting sqref=\"{junk}\"><cfRule type=\"expression\" priority=\"1\"><formula>1</formula></cfRule></conditionalFormatting>"))
            + "</worksheet>";
        var bytes = With(("xl/worksheets/sheet1.xml", sheet));
        Assert.True(bytes.Length < 60_000);
        var limits = new PackageOptions { SourceLimits = new ReaderOptions { MaxCellTextLength = 2_000_000 } }; // Room for a tag this long.
        var ex = Timed(8, () =>
        {
            using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true, limits);
            var data = package.Workbook.Worksheets.Open("Data");
            for (var i = 2; i < 50; i++) data.Row(i).WriteCell(i);
            package.Save();
        }, "Filling the template");
        Assert.Null(ex);
    }

    [Theory]
    [InlineData("B2", true)]
    [InlineData("$B$2", true)]
    [InlineData("b2:d10", true)]
    [InlineData(" D10:B2 ", true)]
    [InlineData("XFD1048576", true)]
    [InlineData("A1:B2:C3", false)]
    [InlineData("XFE1", false)]
    [InlineData("A1048577", false)]
    [InlineData("A0", false)]
    [InlineData("AAAA1", false)]
    [InlineData("A", false)]
    [InlineData("1", false)]
    [InlineData("A1:", false)]
    [InlineData("A:A", false)]
    [InlineData("x", false)]
    [InlineData("", false)]
    [InlineData("A99999999999999999999", false)]
    public void TryParse_AgreesWithParse(string text, bool valid)
    {
        Assert.Equal(valid, RangeReference.TryParse(text, out var range));
        if (valid) Assert.Equal(RangeReference.Parse(text).ToString(), range.ToString());
        else Assert.ThrowsAny<ArgumentException>(() => RangeReference.Parse(text));
    }

    [Fact]
    public void AppendingBelowTheData_DoesNotRecountTheSheetForEveryRow()
    {
        // The documented way to append: Cells[Dimension.LastRow + 1, 1]. On a sheet of 150,000 cells each
        // look at Dimension walked all of them.
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            for (var r = 0; r < 50_000; r++) s.Row().WriteCell(r).WriteCell(r * 2).WriteCell(r * 3);
        });
        using var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true);
        var sheet = package.Workbook.Worksheets.Edit("Data");
        Assert.Equal((50_000, 3), sheet.Cells.Dimension);
        Assert.Null(Timed(6, () =>
        {
            for (var i = 0; i < 20_000; i++) sheet.Cells[sheet.Cells.Dimension.LastRow + 1, 1].Value = i;
        }, "Appending 20,000 rows"));
        Assert.Equal((70_000, 3), sheet.Cells.Dimension);

        // Removing cells can shrink the sheet, and the answer must follow.
        sheet.Cells[70_000, 1].Clear();
        Assert.Equal((69_999, 3), sheet.Cells.Dimension);
        sheet.Cells[80_000, 7].Value = "far";
        Assert.Equal((80_000, 7), sheet.Cells.Dimension);
        sheet.Cells[80_000, 7].Clear();
        Assert.Equal((69_999, 3), sheet.Cells.Dimension);
    }

    [Fact]
    public void TemplateStyle_AskedAgainForTheSameCell_DoesNotReadTheSheetAgain()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var money = wb.Styles.Create(o => o.NumberFormat = "#,##0.00");
            var s = wb.Worksheets.Add("Data");
            for (var r = 0; r < 60_000; r++) s.Row().WriteCell(r).WriteCell(r * 1.5, money);
        });
        using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), Stream.Null, leaveOpen: true);
        var sheet = package.Workbook.Worksheets.Open("Data");
        var first = sheet.TemplateStyle("B59000");
        Assert.Null(Timed(4, () =>
        {
            for (var i = 0; i < 2_000; i++) Assert.Same(first, sheet.TemplateStyle("B59000"));
        }, "Asking 2,000 times"));
        Assert.NotSame(first, sheet.TemplateStyle("A59000"));
    }

    // ------------------------------------------------------------------ what callers collect

    [Fact]
    public void DataTableOfRepeatedSharedTexts_IsChargedOncePerText_NotOncePerCell()
    {
        // Text columns that repeat a few values are the commonest thing in a sheet. Each cell points at one
        // shared string; charging the string again for every cell refused sheets that fit easily.
        var texts = Enumerable.Range(0, 10).Select(i => new string((char)('a' + i), 300)).ToArray();
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.Row().WriteCells("A", "B", "C", "D", "E");
            for (var r = 0; r < 40_000; r++) s.Row().WriteCells(texts[r % 10], texts[(r + 1) % 10], texts[(r + 2) % 10], texts[(r + 3) % 10], texts[(r + 4) % 10]);
        });
        // Charged per cell this would be 200,000 x 632 bytes, about 126 MB. The table really holds a few megabytes.
        using var reader = ExcelReader.Open(new MemoryStream(bytes), false, new ReaderOptions { MaxLoadedPartBytes = 32 * 1024 * 1024 });
        var table = reader.ReadDataTable("Data");
        Assert.Equal(40_000, table.Rows.Count);
        Assert.Equal(texts[3], table.Rows[3][0]);

        // Texts written inline are separate strings and are still charged.
        var inline = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.Row().WriteCells("A");
            for (var r = 0; r < 40_000; r++) s.Row().WriteInlineString(texts[r % 10]);
        });
        using var strict = ExcelReader.Open(new MemoryStream(inline), false, new ReaderOptions { MaxLoadedPartBytes = 16 * 1024 * 1024 });
        Assert.Throws<InvalidDataException>(() => strict.ReadDataTable("Data"));
    }

    [Fact]
    public void DataTableOfSparseWideRows_IsRefusedAtTheBudget_NotBuiltAtAThousandTimesTheFile()
    {
        // 3,000 rows, each with one cell in the last column: 15 KB on disk, a table of 49 million slots.
        var rows = string.Concat(Enumerable.Range(1, 3000).Select(r => $"<row r=\"{r}\"><c r=\"XFD{r}\"><v>1</v></c></row>"));
        var bytes = With(("xl/worksheets/sheet1.xml", Sheet(rows)));
        Assert.True(bytes.Length < 40_000);

        using var reader = ExcelReader.Open(new MemoryStream(bytes), false, new ReaderOptions { MaxLoadedPartBytes = 16 * 1024 * 1024 });
        var before = GC.GetAllocatedBytesForCurrentThread();
        var ex = Assert.Throws<InvalidDataException>(() => reader.ReadDataTable("Data"));
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.Contains("DataTable", ex.Message);
        Assert.Contains("MaxLoadedPartBytes", ex.Message);
        Assert.True(allocated < 200L * 1024 * 1024, $"Refusing the table allocated {allocated / 1048576} MB.");

        // The rows themselves still stream.
        Assert.Equal(3000, reader.ReadRows("Data").Count());
    }

    [Fact]
    public async Task DataTable_WithADataRowWiderThanItsHeader_GetsDistinctColumnNames()
    {
        // The header's second text is "C", which is also the letter the third column would be given.
        var rows = "<row r=\"1\"><c t=\"inlineStr\"><is><t>Name</t></is></c><c t=\"inlineStr\"><is><t>C</t></is></c></row><row r=\"2\"><c><v>1</v></c><c><v>2</v></c><c><v>3</v></c></row>";
        var bytes = With(("xl/worksheets/sheet1.xml", Sheet(rows)));
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var table = reader.ReadDataTable("Data");
        Assert.Equal(new[] { "Name", "C", "C2" }, table.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName));
        Assert.Equal(3.0, table.Rows[0][2]);

        using var asyncReader = await ExcelReader.OpenAsync(new MemoryStream(bytes));
        var asyncTable = await asyncReader.ReadDataTableAsync("Data");
        Assert.Equal(new[] { "Name", "C", "C2" }, asyncTable.Columns.Cast<System.Data.DataColumn>().Select(c => c.ColumnName));
    }

    [Fact]
    public void GetDateTime_OnANumberThatIsNotADate_IsAConversionFailure()
    {
        var bytes = With(("xl/worksheets/sheet1.xml", Sheet("<row r=\"1\"><c><v>1E300</v></c><c><v>-1E300</v></c><c><v>3000000</v></c><c><v>46000.5</v></c></row>")));
        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var row = reader.ReadRows("Data").Single();
        for (var c = 0; c < 3; c++) Assert.Throws<FormatException>(() => row.GetDateTime(c));
        Assert.Equal(new DateTime(2025, 12, 9, 12, 0, 0), row.GetDateTime(3));
    }

    // ------------------------------------------------------------------ the zip container

    /// <summary>
    /// Writes a zip by hand, which the framework's own writer will not do: several names pointing at one
    /// compressed block, the same name twice, names no well-behaved tool would write.
    /// </summary>
    private sealed class RawZip
    {
        private readonly MemoryStream _body = new();
        private readonly List<(byte[] Name, uint Crc, uint Packed, uint Unpacked, uint Offset, ushort Method)> _directory = new();
        private static readonly uint[] Table = Enumerable.Range(0, 256).Select(i =>
        {
            var c = (uint)i;
            for (var k = 0; k < 8; k++) c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            return c;
        }).ToArray();

        private static uint Crc(byte[] data)
        {
            var c = 0xFFFFFFFFu;
            foreach (var b in data) c = Table[(c ^ b) & 0xFF] ^ (c >> 8);
            return c ^ 0xFFFFFFFFu;
        }

        /// <summary>Adds an entry and returns its place in the directory, for <see cref="Alias"/>.</summary>
        public int Add(string name, byte[] content)
        {
            var packed = new MemoryStream();
            using (var deflate = new DeflateStream(packed, CompressionLevel.Optimal, leaveOpen: true)) deflate.Write(content, 0, content.Length);
            var nameBytes = Encoding.UTF8.GetBytes(name);
            var offset = (uint)_body.Position;
            var crc = Crc(content);
            var w = new BinaryWriter(_body, Encoding.UTF8, leaveOpen: true);
            w.Write(0x04034b50u); w.Write((ushort)20); w.Write((ushort)0x0800); w.Write((ushort)8); w.Write((ushort)0); w.Write((ushort)0x21);
            w.Write(crc); w.Write((uint)packed.Length); w.Write((uint)content.Length); w.Write((ushort)nameBytes.Length); w.Write((ushort)0);
            w.Write(nameBytes);
            w.Write(packed.GetBuffer(), 0, (int)packed.Length);
            w.Flush();
            _directory.Add((nameBytes, crc, (uint)packed.Length, (uint)content.Length, offset, 8));
            return _directory.Count - 1;
        }

        /// <summary>Adds another directory record, under another name, for an entry that is already there.</summary>
        public void Alias(int entry, string name)
        {
            var e = _directory[entry];
            _directory.Add((Encoding.UTF8.GetBytes(name), e.Crc, e.Packed, e.Unpacked, e.Offset, e.Method));
        }

        public byte[] ToArray()
        {
            var output = new MemoryStream();
            _body.WriteTo(output);
            var start = (uint)output.Position;
            var w = new BinaryWriter(output, Encoding.UTF8, leaveOpen: true);
            foreach (var e in _directory)
            {
                w.Write(0x02014b50u); w.Write((ushort)20); w.Write((ushort)20); w.Write((ushort)0x0800); w.Write(e.Method); w.Write((ushort)0); w.Write((ushort)0x21);
                w.Write(e.Crc); w.Write(e.Packed); w.Write(e.Unpacked); w.Write((ushort)e.Name.Length); w.Write((ushort)0); w.Write((ushort)0);
                w.Write((ushort)0); w.Write((ushort)0); w.Write(0u); w.Write(e.Offset);
                w.Write(e.Name);
            }
            var size = (uint)output.Position - start;
            w.Write(0x06054b50u); w.Write((ushort)0); w.Write((ushort)0); w.Write((ushort)_directory.Count); w.Write((ushort)_directory.Count);
            w.Write(size); w.Write(start); w.Write((ushort)0);
            w.Flush();
            return output.ToArray();
        }

        /// <summary>A hand-written zip holding the parts of the seed workbook, ready for more to be added.</summary>
        public static RawZip FromSeed()
        {
            var zip = new RawZip();
            using var seed = new ZipArchive(new MemoryStream(Seed()), ZipArchiveMode.Read);
            foreach (var entry in seed.Entries)
            {
                using var content = new MemoryStream();
                using (var source = entry.Open()) source.CopyTo(content);
                zip.Add(entry.FullName, content.ToArray());
            }
            return zip;
        }
    }

    [Fact]
    public void ManyNamesForOneCompressedBlock_DoNotMakeSaveRepackGigabytes()
    {
        // 16 MB of zeros compresses to a few kilobytes; forty directory records point at it under different
        // names. The file is tiny, and copying every "part" would unpack and repack 640 MB. The directory
        // may hold 65,535 records of up to 4 GB each.
        var zip = RawZip.FromSeed();
        var blob = zip.Add("xl/media/b0.bin", new byte[16 * 1024 * 1024]);
        for (var i = 1; i < 40; i++) zip.Alias(blob, $"xl/media/b{i}.bin");
        var bytes = zip.ToArray();
        Assert.True(bytes.Length < 100_000, $"The file is {bytes.Length:N0} bytes.");

        var options = new PackageOptions { SourceLimits = new ReaderOptions { MaxStreamedPartBytes = 64 * 1024 * 1024 } };
        var ex = Timed(20, () => CopyAndSave(bytes, options), "Saving");
        Assert.IsType<InvalidDataException>(ex);
        Assert.Contains("MaxStreamedPartBytes", ex!.Message);

        // The reader opens only the parts the workbook refers to, so it never looks at them.
        Assert.Null(Record.Exception(() => ReadRows(bytes)));
    }

    [Theory]
    [InlineData("../../evil.txt")]
    [InlineData("/abs/evil.txt")]
    [InlineData("xl/../../evil.txt")]
    [InlineData("xl\\media\\evil.txt")]
    [InlineData("xl//double.txt")]
    [InlineData("xl/./dot.txt")]
    [InlineData("xl/media/bell\u0007.txt")]
    public void PartNameThatCouldEscapeTheFolderItIsUnpackedInto_IsNotCopiedToTheOutput(string name)
    {
        var zip = RawZip.FromSeed();
        zip.Add(name, Encoding.UTF8.GetBytes("payload"));
        var bytes = zip.ToArray();

        var ex = Record.Exception(() => CopyAndSave(bytes));
        Assert.IsType<InvalidDataException>(ex);
        Assert.Contains("not a valid part name", ex!.Message);
        Assert.IsType<InvalidDataException>(Record.Exception(() => EditAndSave(bytes)));
    }

    [Fact]
    public void TwoPartsWithOneName_AreRefused()
    {
        var zip = RawZip.FromSeed();
        zip.Add("xl/media/image1.png", new byte[] { 1 });
        zip.Add("xl/media/image1.png", new byte[] { 2 });
        var ex = Record.Exception(() => CopyAndSave(zip.ToArray()));
        Assert.IsType<InvalidDataException>(ex);
        Assert.Contains("more than one part named", ex!.Message);
    }

    [Theory]
    [InlineData("xl/media/image 1 ñ.png")]
    [InlineData("customXml/item1.xml")]
    [InlineData("xl/printerSettings/printerSettings1.bin")]
    [InlineData("[trash]/0000.dat")]
    public void OrdinaryPartNames_AreCopied(string name)
    {
        var zip = RawZip.FromSeed();
        zip.Add(name, Encoding.UTF8.GetBytes("payload"));
        var output = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(zip.ToArray()), output, leaveOpen: true)) package.Save();
        using var result = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        Assert.NotNull(result.GetEntry(name));
    }

    // ------------------------------------------------------------------ damaged structure

    [Theory]
    [InlineData("/")]
    [InlineData(".")]
    [InlineData("worksheets/no-such-sheet.xml")]
    public void SheetWhosePartCannotBeFound_IsReportedAsDamage_AndLeavesThePackageUsable(string target)
    {
        var workbook = $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Ghost\" sheetId=\"2\" r:id=\"rId9\"/></sheets></workbook>";
        var rels = $"<Relationships xmlns=\"{PackageRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId9\" Type=\"{Rel}/worksheet\" Target=\"{target}\"/><Relationship Id=\"rId2\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"{Rel}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>";
        var bytes = With(("xl/workbook.xml", workbook), ("xl/_rels/workbook.xml.rels", rels));

        var output = new MemoryStream();
        using var package = ExcelPackage.OpenTemplate(new MemoryStream(bytes), output, leaveOpen: true);
        Assert.Throws<InvalidDataException>(() => package.Workbook.Worksheets.Open("Ghost"));
        // Nothing was opened in the output, so the good sheet can still be filled and the package saved.
        package.Workbook.Worksheets.Open("Data").Row(5).WriteCell("still works");
        package.Save();
        using var reader = ExcelReader.Open(new MemoryStream(output.ToArray()));
        Assert.Contains(reader.ReadRows("Data"), row => Equals(row[0], "still works"));
    }

    [Fact]
    public void TwoSheetsWithOneName_OrOnePart_AreReportedAsDamage()
    {
        var rels = $"<Relationships xmlns=\"{PackageRel}\"><Relationship Id=\"rId1\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rId2\" Type=\"{Rel}/styles\" Target=\"styles.xml\"/><Relationship Id=\"rId3\" Type=\"{Rel}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>";
        var sameName = $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"DATA\" sheetId=\"2\" r:id=\"rId1\"/></sheets></workbook>";
        var samePart = $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets><sheet name=\"Data\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"Other\" sheetId=\"2\" r:id=\"rId1\"/></sheets></workbook>";
        foreach (var workbook in new[] { sameName, samePart })
        {
            var bytes = With(("xl/workbook.xml", workbook), ("xl/_rels/workbook.xml.rels", rels));
            Assert.IsType<InvalidDataException>(Record.Exception(() => CopyAndSave(bytes)));
            Assert.IsType<InvalidDataException>(Record.Exception(() => EditAndSave(bytes)));
        }
    }

    [Theory]
    [InlineData("xl/workbook.xml", "<workbook xmlns=\"" + Main + "\"/>")]
    [InlineData("xl/styles.xml", "<styleSheet xmlns=\"" + Main + "\"/>")]
    public void PartThatIsOnlyAnEmptyRootElement_IsReportedAsDamage(string part, string content)
    {
        var bytes = With((part, content));
        Assert.IsType<InvalidDataException>(Record.Exception(() => CopyAndSave(bytes)));
    }

    [Fact]
    public void PartWhoseRootIsTheElementThatWouldBeDropped_IsReportedAsDamage()
    {
        // Found by fuzzing a workbook Excel wrote: the content types part starts with the calculation chain's
        // own entry instead of the list that should hold it.
        var source = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("Data").Row().WriteCell(1).WriteFormula("A1*2", 2.0));
        var output = new MemoryStream();
        using (var from = new ZipArchive(new MemoryStream(source), ZipArchiveMode.Read))
        using (var to = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in from.Entries)
            {
                using var target = to.CreateEntry(entry.FullName).Open();
                if (entry.FullName == "[Content_Types].xml")
                {
                    using var writer = new StreamWriter(target);
                    writer.Write("<Override xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\" PartName=\"/xl/calcChain.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml\"/>");
                    continue;
                }
                using var content = entry.Open();
                content.CopyTo(target);
            }
            // The chain itself and its relationship, so the save has something to drop.
            using (var writer = new StreamWriter(to.CreateEntry("xl/calcChain.xml").Open())) writer.Write($"<calcChain xmlns=\"{Main}\"><c r=\"B1\" i=\"1\"/></calcChain>");
        }
        var bytes = output.ToArray();
        // Point the workbook at the chain.
        string rels;
        using (var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read))
        using (var reader = new StreamReader(zip.GetEntry("xl/_rels/workbook.xml.rels")!.Open())) rels = reader.ReadToEnd();
        var patched = new MemoryStream();
        using (var from = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read))
        using (var to = new ZipArchive(patched, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in from.Entries)
            {
                using var target = to.CreateEntry(entry.FullName).Open();
                if (entry.FullName == "xl/_rels/workbook.xml.rels")
                {
                    using var writer = new StreamWriter(target);
                    writer.Write(rels.Replace("</Relationships>", $"<Relationship Id=\"rIdCalc\" Type=\"{Rel}/calcChain\" Target=\"calcChain.xml\"/></Relationships>"));
                    continue;
                }
                using var content = entry.Open();
                content.CopyTo(target);
            }
        }
        Assert.IsType<InvalidDataException>(Record.Exception(() => CopyAndSave(patched.ToArray())));
    }

    [Fact]
    public void SharedStringNestedInAnExtension_DoesNotShiftTheStringsThatFollow()
    {
        // An "si" inside an extension element is not an item of the table. Counting it made every string
        // added afterwards point one place too far.
        var sst = $"<sst xmlns=\"{Main}\"><si><t>Name</t></si><si><t>Qty</t></si><si><t>a</t></si><extLst><ext uri=\"x\"><si><t>not an item</t></si></ext></extLst></sst>";
        var bytes = With(("xl/sharedStrings.xml", sst));
        var output = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(bytes), output, leaveOpen: true))
        {
            var sheet = package.Workbook.Worksheets.Edit("Data");
            sheet.Cells["A3"].Value = "appended one";
            sheet.Cells["A4"].Value = "appended two";
            package.Save();
        }
        using var reader = ExcelReader.Open(new MemoryStream(output.ToArray()));
        var rows = reader.ReadRows("Data").ToList();
        Assert.Equal("Name", rows[0][0]);
        Assert.Equal("a", rows[1][0]);
        Assert.Equal("appended one", rows[2][0]);
        Assert.Equal("appended two", rows[3][0]);
    }

    [Fact]
    public void NumberFormatIdAtTheTopOfTheIntegerRange_DoesNotMakeTheNextIdOverflow()
    {
        var styles = $"<styleSheet xmlns=\"{Main}\"><numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"0.0\"/><numFmt numFmtId=\"2147483647\" formatCode=\"0.00\"/></numFmts><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"1\"><xf numFmtId=\"0\"/></cellXfs></styleSheet>";
        var bytes = With(("xl/styles.xml", styles));
        var output = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(bytes), output, leaveOpen: true))
        {
            var custom = package.Workbook.Styles.Create(o => o.NumberFormat = "0.000 \"kg\"");
            var cell = package.Workbook.Worksheets.Edit("Data").Cells["D4"];
            cell.Value = 1.5;
            cell.Style = custom;
            package.Save();
        }
        string written;
        using (var zip = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read))
        using (var reader = new StreamReader(zip.GetEntry("xl/styles.xml")!.Open())) written = reader.ReadToEnd();
        Assert.DoesNotContain("numFmtId=\"-", written);
        Assert.Contains("numFmtId=\"165\" formatCode=\"0.000 &quot;kg&quot;\"", written);
    }

    [Fact]
    public void TwoFormatIdsWithOneCode_AreBothDatesWhenEdited()
    {
        var styles = $"<styleSheet xmlns=\"{Main}\"><numFmts count=\"2\"><numFmt numFmtId=\"164\" formatCode=\"yyyy-mm-dd\"/><numFmt numFmtId=\"165\" formatCode=\"yyyy-mm-dd\"/></numFmts><fonts count=\"1\"><font/></fonts><fills count=\"1\"><fill/></fills><borders count=\"1\"><border/></borders><cellXfs count=\"3\"><xf numFmtId=\"0\"/><xf numFmtId=\"164\"/><xf numFmtId=\"165\"/></cellXfs></styleSheet>";
        var bytes = With(("xl/styles.xml", styles), ("xl/worksheets/sheet1.xml", Sheet("<row r=\"1\"><c r=\"A1\" s=\"1\"><v>46000</v></c><c r=\"B1\" s=\"2\"><v>46000</v></c></row>")));
        using var package = ExcelPackage.Edit(new MemoryStream(bytes), Stream.Null, leaveOpen: true);
        var sheet = package.Workbook.Worksheets.Edit("Data");
        Assert.IsType<DateTime>(sheet.Cells["A1"].Value);
        Assert.IsType<DateTime>(sheet.Cells["B1"].Value);
    }

    // ------------------------------------------------------------------ formulas from a file

    [Theory]
    [InlineData(3000, "(")]      // Nested parentheses.
    [InlineData(100_000, "-")]   // A long run of signs.
    [InlineData(3000, "ABS(")]   // Nested function calls.
    public void DeeplyNestedFormula_IsLeftUncached_NotFollowedUntilTheStackOverflows(int depth, string open)
    {
        // On a thread with a small stack, as a request thread has: a stack overflow cannot be caught and
        // would end the process, taking this test run with it.
        var close = open == "-" ? "" : ")";
        var formula = string.Concat(Enumerable.Repeat(open, depth)) + "1" + string.Concat(Enumerable.Repeat(close, depth));
        Exception? failure = null;
        byte[]? bytes = null;
        var thread = new Thread(() =>
        {
            try
            {
                bytes = ExcelPackage.CreateBytes(wb =>
                {
                    var sheet = wb.Worksheets.Add("Data");
                    sheet.CacheFormulaResults();
                    sheet.Row().WriteCell(1).WriteFormula(formula).WriteFormula("((A1+1)*2)");
                });
            }
            catch (Exception ex) { failure = ex; }
        }, 256 * 1024);
        thread.Start();
        thread.Join();

        Assert.Null(failure);
        using var reader = ExcelReader.Open(new MemoryStream(bytes!));
        var row = reader.ReadRows("Data").Single();
        Assert.Null(row[1]);          // No cached value: the formula is written, Excel will calculate it.
        Assert.Equal(4.0, row[2]);    // An ordinary formula on the same row is still cached.
    }

    [Fact]
    public void FormulaValueStore_StopsAtItsCeiling_AndThenReportsEverythingUnknown()
    {
        // One value per page, the costliest thing a sheet can make the store do: 8 KB for 8 bytes.
        var store = new FormulaValueStore();
        var before = GC.GetAllocatedBytesForCurrentThread();
        var written = 0;
        for (var column = 1; column <= 16_384 && written < 60_000; column++)
            for (var row = 1; row <= 1_048_576 && written < 60_000; row += 1024, written++)
                store.SetNumber(row, column, 42);
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        // 60,000 pages would be 480 MB; the store gives up at 256 MB.
        Assert.True(allocated < FormulaValueStore.MaxBytes + 64L * 1024 * 1024, $"The store allocated {allocated / 1048576} MB.");
        Assert.Equal(FormulaCellState.Unknown, store.Get(1, 1, out _));
        Assert.Equal(FormulaCellState.Unknown, store.Get(500_000, 7, out _));
        store.SetNumber(5, 5, 1); // Still safe to call.
        Assert.Equal(FormulaCellState.Unknown, store.Get(5, 5, out _));

        // A dense sheet is nowhere near the ceiling: a million numbers take 8 MB.
        var dense = new FormulaValueStore();
        for (var row = 1; row <= 1_000_000; row++) dense.SetNumber(row, 1, row);
        Assert.Equal(FormulaCellState.Number, dense.Get(999_999, 1, out var value));
        Assert.Equal(999_999, value);
    }

    // ------------------------------------------------------------------ saving

    [Fact]
    public void SaveThatCannotPutTheFileInPlace_CanBeRetried_AndLosesNothing()
    {
        // The obstacle is a folder sitting where the file should go, which stops the move on every operating
        // system. (A file held open stops it on Windows only; that case has its own test below.)
        var dir = Path.Combine(Path.GetTempPath(), "ss-retry-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, "report.xlsx");
            Directory.CreateDirectory(path);

            using var package = ExcelPackage.Create(path);
            package.Workbook.Worksheets.Add("Data").Row().WriteCell("new content");

            var ex = Assert.Throws<IOException>(() => package.Save());
            Assert.Contains("call Save again", ex.Message);
            Assert.False(package.IsSaved);
            Assert.Throws<IOException>(() => package.Save()); // Still in the way: the same answer, not a silent success.
            Assert.False(package.IsSaved);

            // The obstacle goes; the same package saves, and the workbook it built is the one that lands.
            Directory.Delete(path);
            package.Save();
            Assert.True(package.IsSaved);
            using (var reader = ExcelReader.Open(path)) Assert.Equal("new content", reader.ReadRows(0).First()[0]);
            package.Dispose();
            Assert.Single(Directory.GetFileSystemEntries(dir));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void SaveThatNeverGetsTheFileInPlace_LeavesNoTemporaryFile()
    {
        var dir = Path.Combine(Path.GetTempPath(), "ss-retry-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, "report.xlsx");
            Directory.CreateDirectory(path);
            using (var package = ExcelPackage.Create(path))
            {
                package.Workbook.Worksheets.Add("Data").Row().WriteCell("new content");
                Assert.Throws<IOException>(() => package.Save());
                Assert.Single(Directory.GetFiles(dir)); // The finished workbook, waiting in its temporary file.
            }
            Assert.Empty(Directory.GetFiles(dir));
            Assert.True(Directory.Exists(path));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void SaveOverAFileSomeoneHasOpen_LeavesItUntouched_AndSucceedsOnceItIsClosed()
    {
        // Windows refuses to replace a file another handle holds open, which is what Excel does to a workbook
        // it is showing. Other systems let the replace through, so there is nothing to retry there.
        if (!OperatingSystem.IsWindows()) return;

        var dir = Path.Combine(Path.GetTempPath(), "ss-retry-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, "report.xlsx");
            File.WriteAllText(path, "ORIGINAL");

            using var package = ExcelPackage.Create(path);
            package.Workbook.Worksheets.Add("Data").Row().WriteCell("new content");
            using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None))
            {
                Assert.Throws<IOException>(() => package.Save());
                Assert.False(package.IsSaved);
            }
            Assert.Equal("ORIGINAL", File.ReadAllText(path));

            package.Save();
            Assert.True(package.IsSaved);
            using (var reader = ExcelReader.Open(path)) Assert.Equal("new content", reader.ReadRows(0).First()[0]);
            package.Dispose();
            Assert.Single(Directory.GetFiles(dir));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }
    [Fact]
    public void TargetWithALongFileName_StillWorks()
    {
        // The temporary file used to be named after the target plus 37 characters, which pushed a long but
        // valid name over the limit for a file name.
        var dir = Path.Combine(Path.GetTempPath(), "ss-long-" + Guid.NewGuid().ToString("N").Substring(0, 8));
        Directory.CreateDirectory(dir);
        try
        {
            var path = Path.Combine(dir, new string('n', 240) + ".xlsx");
            using (var package = ExcelPackage.Create(path))
            {
                package.Workbook.Worksheets.Add("Data").Row().WriteCell(1);
                package.Save();
            }
            Assert.True(File.Exists(path));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void ReaderThatFailsToOpen_ReleasesTheStreamItWasGiven()
    {
        var notAZip = new TrackingStream(Encoding.UTF8.GetBytes("this is not a workbook"));
        Assert.Throws<InvalidDataException>(() => ExcelReader.Open(notAZip, leaveOpen: false));
        Assert.True(notAZip.Disposed);

        var noWorkbook = new MemoryStream();
        using (var zip = new ZipArchive(noWorkbook, ZipArchiveMode.Create, leaveOpen: true)) zip.CreateEntry("readme.txt");
        var tracked = new TrackingStream(noWorkbook.ToArray());
        Assert.Throws<InvalidDataException>(() => ExcelReader.Open(tracked, leaveOpen: false));
        Assert.True(tracked.Disposed);

        var kept = new TrackingStream(Encoding.UTF8.GetBytes("this is not a workbook"));
        Assert.Throws<InvalidDataException>(() => ExcelReader.Open(kept, leaveOpen: true));
        Assert.False(kept.Disposed);

        var output = new TrackingStream(Array.Empty<byte>(), writable: true);
        Assert.Throws<InvalidDataException>(() => ExcelPackage.Edit(new MemoryStream(Encoding.UTF8.GetBytes("garbage")), output, leaveOpen: false));
        Assert.True(output.Disposed);
    }

    [Fact]
    public async Task SourceThatFailsWhileItIsCopied_IsReleased()
    {
        // A stream that cannot seek (an upload) is copied before it is read. If it breaks half-way, the
        // caller who handed it over still expects it to be closed.
        var sync = new FailingStream();
        Assert.Throws<IOException>(() => ExcelReader.Open(sync, leaveOpen: false));
        Assert.True(sync.Disposed);

        var async = new FailingStream();
        await Assert.ThrowsAsync<IOException>(() => ExcelReader.OpenAsync(async, leaveOpen: false));
        Assert.True(async.Disposed);

        var kept = new FailingStream();
        Assert.Throws<IOException>(() => ExcelReader.Open(kept, leaveOpen: true));
        Assert.False(kept.Disposed);
    }

    /// <summary>A stream that cannot seek and breaks after a few kilobytes, like a dropped upload.</summary>
    private sealed class FailingStream : Stream
    {
        private int _served;
        public bool Disposed { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_served > 8000) throw new IOException("The connection was reset.");
            _served += count;
            return count;
        }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }

    private sealed class TrackingStream : MemoryStream
    {
        public TrackingStream(byte[] bytes, bool writable = false) : base()
        {
            Write(bytes, 0, bytes.Length);
            Position = 0;
            _ = writable;
        }

        public bool Disposed { get; private set; }

        protected override void Dispose(bool disposing)
        {
            Disposed = true;
            base.Dispose(disposing);
        }
    }
}
