using System.IO.Compression;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Internal;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// <see cref="PackageOptions.StringStorage"/>: where the text of a cell is stored (the shared string table or the cell
/// itself) must never change what a reader sees, and must follow the rules in the documentation: a sheet small
/// enough to fit in the first sample is stored exactly as the shared table stores it, a column of unique text goes
/// inline, a column that repeats long or varied text is shared, and memory stays bounded.
/// </summary>
public class StringStorageTests
{
    private static PackageOptions Options(StringStorage storage, bool references = false) => new PackageOptions
    {
        StringStorage = storage,
        WriteCellReferences = references,
        CompressionLevel = CompressionLevel.Fastest,
    };

    /// <summary>A repeatable, evenly spread pick from count values, so repeats are at random distances (a plain multiple would visit each value once before any came back).</summary>
    private static int Pick(int i, int count)
    {
        if (_picks.Length <= i)
        {
            var random = new Random(20261116); // A fixed seed: never string.GetHashCode, which differs from one process to the next.
            // Grown in one step to a size no test here outgrows; a table that grew one entry at a time would be rebuilt for every cell.
            _picks = Enumerable.Range(0, Math.Max(i + 1, 300_000)).Select(_ => random.Next()).ToArray();
        }
        return _picks[i] % count;
    }

    private static int[] _picks = Array.Empty<int>();

    private static string Part(byte[] workbook, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(workbook), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(name)!.Open(), new UTF8Encoding(false));
        return reader.ReadToEnd();
    }

    private static int Count(string text, string pattern)
    {
        var count = 0;
        for (var at = text.IndexOf(pattern, StringComparison.Ordinal); at >= 0; at = text.IndexOf(pattern, at + pattern.Length, StringComparison.Ordinal)) count++;
        return count;
    }

    private static List<ExcelRow> Read(byte[] workbook)
    {
        using var reader = ExcelReader.Open(new MemoryStream(workbook), false);
        return reader.ReadRows(0).ToList();
    }

    private static void AssertReadsBack(byte[] workbook, IReadOnlyList<string?[]> expected)
    {
        var rows = Read(workbook);
        Assert.Equal(expected.Count, rows.Count);
        for (var r = 0; r < expected.Count; r++)
            for (var c = 0; c < expected[r].Length; c++)
                Assert.Equal(expected[r][c], rows[r][c] as string);
    }

    // ---- the options -----------------------------------------------------------------------------------------------

    [Fact]
    public void TheDefaultIsAuto_AndTheOldSwitchStillMeansInline()
    {
        var options = new PackageOptions();
        Assert.Equal(StringStorage.Auto, options.StringStorage);
        Assert.True(options.UseSharedStrings);

        // UseSharedStrings = false wins over every storage mode, as it always did.
        foreach (var storage in new[] { StringStorage.Auto, StringStorage.Shared, StringStorage.Inline })
        {
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("S");
                for (var i = 0; i < 20; i++) s.Row().WriteCell("text " + i % 3);
                Assert.Equal(0, wb.SharedStrings.UniqueCount);
            }, new PackageOptions { StringStorage = storage, UseSharedStrings = false });
            var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
            Assert.Equal(20, Count(sheet, "t=\"inlineStr\""));
            Assert.Equal(0, Count(sheet, "t=\"s\""));
        }
    }

    [Fact]
    public void Shared_PutsEveryTextInTheTable_AndInline_PutsNoneThere()
    {
        var rows = Enumerable.Range(0, 3000).Select(i => new[] { "unique " + i, "repeated " + i % 4 }).ToList();

        var shared = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            foreach (var r in rows) s.Row().WriteCell(r[0]).WriteCell(r[1]);
            Assert.Equal(3004, wb.SharedStrings.UniqueCount);
        }, Options(StringStorage.Shared));
        Assert.Equal(0, Count(Part(shared, "xl/worksheets/sheet1.xml"), "inlineStr"));
        AssertReadsBack(shared, rows);

        var inline = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            foreach (var r in rows) s.Row().WriteCell(r[0]).WriteCell(r[1]);
            Assert.Equal(0, wb.SharedStrings.UniqueCount);
        }, Options(StringStorage.Inline));
        Assert.Equal(0, Count(Part(inline, "xl/worksheets/sheet1.xml"), "t=\"s\""));
        AssertReadsBack(inline, rows);
    }

    // ---- Auto ------------------------------------------------------------------------------------------------------

    /// <summary>A sheet with fewer rows than the shortest sample is stored exactly as <see cref="StringStorage.Shared"/> stores it.</summary>
    [Fact]
    public void Auto_ASheetWithFewerRowsThanTheShortestSample_IsStoredExactlyAsShared()
    {
        void Fill(Workbook wb)
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < StringColumnPolicy.EarlyExitCells - 1; i++)
                s.Row().WriteCell("id " + i).WriteCell(i % 5 == 0 ? "five" : "other").WriteCell(i).WriteCell(i % 2 == 0 ? "a & b" : " padded ");
        }

        var auto = ExcelPackage.CreateBytes(Fill, Options(StringStorage.Auto));
        var shared = ExcelPackage.CreateBytes(Fill, Options(StringStorage.Shared));
        Assert.Equal(Part(shared, "xl/worksheets/sheet1.xml"), Part(auto, "xl/worksheets/sheet1.xml"));
        Assert.Equal(Part(shared, "xl/sharedStrings.xml"), Part(auto, "xl/sharedStrings.xml"));
    }

    [Fact]
    public void Auto_AColumnOfUniqueText_GoesInline_AndTheTableStaysSmall()
    {
        const int rows = 40_000;
        long uniqueInTable = 0;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell("Order SO-" + i.ToString("D8")).WriteCell("customer" + i + "@example.com");
            uniqueInTable = wb.SharedStrings.UniqueCount;
        }, Options(StringStorage.Auto));

        // Only the samples (the first window and a few later ones) of the two columns went through the table.
        Assert.True(uniqueInTable < 8 * StringColumnPolicy.ProbeCells, $"The table holds {uniqueInTable} texts.");
        var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"inlineStr\"") > 2 * rows - 8 * StringColumnPolicy.ProbeCells);
        AssertReadsBack(bytes, Enumerable.Range(0, rows).Select(i => new[] { "Order SO-" + i.ToString("D8"), "customer" + i + "@example.com" }).ToList());
    }

    // ---- a second sheet is judged by what it does, not by what an earlier sheet put in the table ----------------------

    /// <summary>The same unique texts on two sheets, as in a workbook that lists the same ids twice.</summary>
    [Theory]
    [InlineData(20_000)]
    [InlineData(100_000)]
    public void Auto_AUniqueColumnOnASecondSheet_IsNotJudgedByTheFirstSheetsEntries(int rows)
    {
        long OneSheet()
        {
            long unique = 0;
            ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("S");
                for (var i = 0; i < rows; i++) s.Row().WriteCell("Order SO-" + i.ToString("D8"));
                unique = wb.SharedStrings.UniqueCount;
            }, Options(StringStorage.Auto));
            return unique;
        }

        long twoSheets = 0;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            foreach (var name in new[] { "S", "T" })
            {
                var s = wb.Worksheets.Add(name);
                for (var i = 0; i < rows; i++) s.Row().WriteCell("Order SO-" + i.ToString("D8"));
            }
            twoSheets = wb.SharedStrings.UniqueCount;
        }, Options(StringStorage.Auto));

        var one = OneSheet();
        // Before the fix the second sheet read the first sheet's entries as repeats and shared its whole column.
        Assert.Equal(one, twoSheets);
        Assert.True(Count(Part(bytes, "xl/worksheets/sheet2.xml"), "t=\"inlineStr\"") > rows - 8 * StringColumnPolicy.ProbeCells);
    }

    /// <summary>
    /// Every sheet takes a scope, which needs an array as long as the table is then. Sized exactly it is reallocated for every sheet
    /// (quadratic in the number of sheets: 1,000 small sheets cost three times the memory they should); it must grow geometrically.
    /// </summary>
    [Fact]
    public void Auto_ManySmallSheets_DoNotReallocateTheScopeArrayForEverySheet()
    {
        // The table itself: 300 scopes, each begun after 640 more entries (what a 200-row sheet of five unique columns adds).
        var table = new SharedStringTable(long.MaxValue / 4);
        long allocatedByScopes = 0;
        for (var sheet = 0; sheet < 300; sheet++)
        {
            for (var i = 0; i < 640; i++) table.TryGetOrAdd("s" + sheet + "_" + i, out _);
            var before = GC.GetAllocatedBytesForCurrentThread();
            table.BeginScope(out _);
            allocatedByScopes += GC.GetAllocatedBytesForCurrentThread() - before;
        }
        // 192,000 entries: four bytes each is 0.8 MB; doubling costs at most about three times that. Resizing to the exact size every time cost 115 MB.
        Assert.True(allocatedByScopes < 4_000_000, $"The scopes allocated {allocatedByScopes:N0} bytes.");

        // And through a package: the first sheet is the same bytes whether or not other sheets follow it, and every sheet reads back.
        void Fill(Workbook wb, int sheets)
        {
            for (var n = 0; n < sheets; n++)
            {
                var s = wb.Worksheets.Add("S" + n);
                for (var i = 0; i < 200; i++) s.Row().WriteCell("a" + n + "_" + i).WriteCell("b" + n + "_" + i).WriteCell("c" + n + "_" + i).WriteCell("d" + n + "_" + i).WriteCell("e" + n + "_" + i);
            }
        }
        var one = ExcelPackage.CreateBytes(wb => Fill(wb, 1), Options(StringStorage.Auto));
        var many = ExcelPackage.CreateBytes(wb => Fill(wb, 300), Options(StringStorage.Auto));
        Assert.Equal(Part(one, "xl/worksheets/sheet1.xml"), Part(many, "xl/worksheets/sheet1.xml"));
        using var reader = ExcelReader.Open(new MemoryStream(many), false);
        Assert.Equal(300, reader.SheetNames.Count);
        Assert.Equal("e299_199", reader.ReadRows(299).Last()[4]);
    }

    /// <summary>A column that really repeats must still be shared on a second sheet, with its own values and with the first sheet's.</summary>
    [Fact]
    public void Auto_ARepeatingColumnOnASecondSheet_IsStillShared()
    {
        var phrases = Enumerable.Range(0, 300).Select(i => "Waiting for the warehouse to confirm stock, batch " + i).ToArray();
        var other = Enumerable.Range(0, 300).Select(i => "Customer asked for a call back about order " + i).ToArray();
        const int rows = 20_000;
        long unique = 0;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell("Order SO-" + i.ToString("D8")).WriteCell(phrases[Pick(i, 300)]);
            var t = wb.Worksheets.Add("T");
            // Column A: fresh repeating text. Column B: the repeating text the first sheet already put in the table.
            for (var i = 0; i < rows; i++) t.Row().WriteCell(other[Pick(i, 300)]).WriteCell(phrases[Pick(i + 1, 300)]);
            unique = wb.SharedStrings.UniqueCount;
        }, Options(StringStorage.Auto));

        var sheet2 = Part(bytes, "xl/worksheets/sheet2.xml");
        Assert.True(Count(sheet2, "t=\"s\"") >= 2 * rows - 2 * StringColumnPolicy.ProbeCells, "Both repeating columns of the second sheet go through the table.");
        // The 300 + 300 repeating texts and the first sheet's sampled ids (the early exit, then the two second looks); nothing more.
        var sampledIds = StringColumnPolicy.EarlyExitCells + 2 * StringColumnPolicy.LookCells;
        Assert.True(unique <= 600 + sampledIds + 16, $"The table holds {unique} texts.");
    }

    [Fact]
    public void Auto_AColumnThatRepeatsLongText_IsShared()
    {
        var phrases = new[] { "Pending approval by the finance team", "Waiting for the warehouse to confirm stock", "Customer asked for a call back", "Shipped, signature required" };
        const int rows = 20_000;
        long uniqueInTable = 0;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell(phrases[i * 7 % 4]).WriteCell("id" + i);
            uniqueInTable = wb.SharedStrings.UniqueCount;
        }, Options(StringStorage.Auto));

        var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"s\"") >= rows, "Every cell of the repeating column goes through the table.");
        Assert.True(uniqueInTable < 8 * StringColumnPolicy.ProbeCells + 4);
        AssertReadsBack(bytes, Enumerable.Range(0, rows).Select(i => new[] { phrases[i * 7 % 4], "id" + i }).ToList());
    }

    /// <summary>Pins the rule that short texts from a small set are not shared: the zip compressor folds them away, and sharing them costs time.</summary>
    [Fact]
    public void Auto_AColumnOfFewShortValues_GoesInlineAfterTheFirstSample()
    {
        var statuses = new[] { "New", "Pending", "Shipped", "Delivered", "Cancelled" };
        const int rows = 20_000;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell(statuses[i * 3 % 5]);
        }, Options(StringStorage.Auto));
        var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"inlineStr\"") > rows - 8 * StringColumnPolicy.ProbeCells);
        AssertReadsBack(bytes, Enumerable.Range(0, rows).Select(i => new[] { statuses[i * 3 % 5] }).ToList());
    }

    [Fact]
    public void Auto_AColumnWithManyDistinctRepeatedValues_IsSharedEvenWhenTheyAreShort()
    {
        const int rows = 30_000;
        long uniqueInTable = 0;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell("N" + Pick(i, 1500));
            uniqueInTable = wb.SharedStrings.UniqueCount;
        }, Options(StringStorage.Auto));
        Assert.Equal(1500, uniqueInTable);
        Assert.Equal(0, Count(Part(bytes, "xl/worksheets/sheet1.xml"), "inlineStr"));
        AssertReadsBack(bytes, Enumerable.Range(0, rows).Select(i => new[] { "N" + Pick(i, 1500) }).ToList());
    }

    /// <summary>Columns that change their nature half way down the sheet, in both directions, and ones that mix: every cell must still read back exactly.</summary>
    [Fact]
    public void Auto_ColumnsThatChangeTheirBehaviourMidSheet_StillReadBackExactly()
    {
        const int rows = 150_000;
        const int change = 20_000;
        var phrases = new[] { "Pending approval by the finance team", "Waiting for the warehouse to confirm stock", "Customer asked for a call back" };
        string?[] Row(int i) => new string?[]
        {
            // repeats for the first stretch, then every value is new
            i < change ? phrases[i % 3] : "free text number " + i,
            // new for the first stretch, then repeats
            i < change ? "ref " + i : phrases[i % 3],
            // a short, repeating value all the way
            "S" + i % 4,
            // half new, half repeated, in blocks
            i / 700 % 2 == 0 ? "block " + i : phrases[i % 3],
            // with characters that need care
            i % 3 == 0 ? "A & B <" + i + "> é€😀 " : " lead" + i + "_x0041_ ",
        };

        var expected = new List<string?[]>();
        long uniqueInTable = 0;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++)
            {
                var values = Row(i);
                expected.Add(values);
                var w = s.Row();
                foreach (var v in values) w.WriteCell(v);
            }
            uniqueInTable = wb.SharedStrings.UniqueCount;
        }, Options(StringStorage.Auto));

        var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"s\"") > 0 && Count(sheet, "t=\"inlineStr\"") > 0, "Both kinds of cell are in the file.");
        Assert.True(uniqueInTable < expected.SelectMany(r => r).Distinct().Count(), "The table did not take every text.");
        AssertReadsBack(bytes, expected);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(doc));
    }

    /// <summary>
    /// Text that has to be escaped (<c>_xHHHH_</c>) is kept in the shared table by Auto even in an inline column, because
    /// openpyxl and pandas decode the escapes in shared strings only. Written inline on request it stays inline.
    /// </summary>
    [Fact]
    public void Auto_TextThatNeedsEscaping_GoesThroughTheTable_EvenInAnInlineColumn()
    {
        string Text(int i) => i % 10 == 0 ? "control \u0001 and literal _x0041_ " + i : "unique " + i;
        const int rows = 3000;
        byte[] Build(StringStorage storage) => ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell(Text(i));
        }, Options(storage));

        var auto = Build(StringStorage.Auto);
        var inline = Build(StringStorage.Inline);
        var expected = Enumerable.Range(0, rows).Select(i => new[] { Text(i) }).ToList();
        AssertReadsBack(auto, expected);
        AssertReadsBack(inline, expected);

        // After the first sample the column is inline, but the escaped texts are still in the table.
        var sheet = Part(auto, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"inlineStr\"") > rows / 2);
        Assert.True(Count(Part(auto, "xl/sharedStrings.xml"), "_x0001_") >= rows / 10 - 1, "Every escaped text is in the table.");
        Assert.Equal(0, Count(sheet, "_x0001_"));
        Assert.True(Count(Part(inline, "xl/worksheets/sheet1.xml"), "_x0001_") >= rows / 10);
    }
    [Fact]
    public void Auto_ACellOfEveryKind_ReadsBackThroughBothStorageModes()
    {
        // Every other API that writes text: hyperlinks, formula results, rich text, and a table whose header is text.
        foreach (var storage in new[] { StringStorage.Auto, StringStorage.Shared, StringStorage.Inline })
        {
            var bytes = ExcelPackage.CreateBytes(wb =>
            {
                var s = wb.Worksheets.Add("S");
                s.AddTable("T", 1, 1, 3001, 4);
                s.Row().WriteCell("Name").WriteCell("Link").WriteCell("Formula").WriteCell("Rich");
                for (var i = 0; i < 3000; i++)
                    s.Row().WriteCell("name " + i % 7)
                        .WriteHyperlink("link " + i, "https://example.com/" + i)
                        .WriteFormula("A" + (i + 2), "name " + i % 7)
                        .WriteRichText(new RichText("bold " + i) { Bold = true }, new RichText(" plain"));
            }, Options(storage));

            var rows = Read(bytes);
            Assert.Equal(3001, rows.Count);
            Assert.Equal("Name", rows[0][0]);
            Assert.Equal("Rich", rows[0][3]);
            Assert.Equal("name 2", rows[3][0]);
            Assert.Equal("link 2", rows[3][1]);
            Assert.Equal("bold 2 plain", rows[3][3]);
            using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
            Assert.Empty(new OpenXmlValidator().Validate(doc));
        }
    }

    [Fact]
    public void Auto_TheMemoryBudgetStillMakesNewTextsInline_AndSharedOnesStayShared()
    {
        var phrases = Enumerable.Range(0, 400).Select(i => "a long phrase that repeats, number " + i).ToArray();
        var rows = Enumerable.Range(0, 30_000).Select(i => new[] { phrases[i * 31 % 400], "unique " + i }).ToList();
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            foreach (var r in rows) s.Row().WriteCell(r[0]).WriteCell(r[1]);
        }, new PackageOptions { StringStorage = StringStorage.Auto, MaxSharedStringChars = 6000, CompressionLevel = CompressionLevel.Fastest });
        AssertReadsBack(bytes, rows);
    }

    [Fact]
    public void Auto_RandomAccessCellsAndAutoFit_ReadBackExactly()
    {
        var expected = new List<string?[]>();
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.AutoFitColumns();
            for (var i = 0; i < 6000; i++)
            {
                var values = new string?[] { "key" + i, "Delivered to the customer", "k" + i % 3 };
                expected.Add(values);
                var w = s.Row();
                foreach (var v in values) w.WriteCell(v);
            }
        }, Options(StringStorage.Auto));
        AssertReadsBack(bytes, expected);

        var random = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 1; i <= 6000; i++)
            {
                s.Cells[i, 3].Value = "Delivered to the customer";
                s.Cells[i, 1].Value = "key" + (i - 1);
            }
        }, Options(StringStorage.Auto));
        var rows = Read(random);
        Assert.Equal(6000, rows.Count);
        Assert.Equal("key5999", rows[5999][0]);
        Assert.Equal("Delivered to the customer", rows[5999][2]);
    }

    [Fact]
    public void Auto_WorkbooksFilledIntoATemplate_KeepTheTemplatesOwnStrings()
    {
        var template = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            s.Row().WriteCell("Header A").WriteCell("Header B");
            s.Row().WriteCell("Header A").WriteCell("Header B");
        }, Options(StringStorage.Shared));

        using var output = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true, Options(StringStorage.Auto)))
        {
            var s = package.Workbook.Worksheets.Open("Data");
            for (var i = 0; i < 12_000; i++) s.Row().WriteCell("Header A").WriteCell("unique " + i).WriteCell("Customer asked for a call back");
            package.Save();
        }

        var rows = Read(output.ToArray());
        Assert.Equal(12_000, rows.Count);
        Assert.Equal("Header A", rows[0][0]);
        Assert.Equal("unique 11999", rows[11_999][1]);
        Assert.Equal("Customer asked for a call back", rows[11_999][2]);
        using var doc = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
        Assert.Empty(new OpenXmlValidator().Validate(doc));
    }

    [Fact]
    public void Auto_EditingAWorkbook_KeepsEveryCell()
    {
        var original = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Data");
            for (var i = 0; i < 3000; i++) s.Row().WriteCell("row " + i).WriteCell("Waiting for the warehouse to confirm stock");
        }, Options(StringStorage.Shared));

        using var output = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(original), output, leaveOpen: true, Options(StringStorage.Auto)))
        {
            var s = package.Workbook.Worksheets.Edit("Data");
            for (var i = 0; i < 3000; i += 10) s.Cells[i + 1, 3].Value = "edited " + i;
            package.Save();
        }
        var rows = Read(output.ToArray());
        Assert.Equal(3000, rows.Count);
        Assert.Equal("row 2999", rows[2999][0]);
        Assert.Equal("Waiting for the warehouse to confirm stock", rows[2999][1]);
        Assert.Equal("edited 2990", rows[2990][2]);
    }

    // ---- the shared strings part ---------------------------------------------------------------------------------

    /// <summary>
    /// The shared strings part is composed in chunks and handed to the XML writer as raw text. It must be what one XmlWriter call per
    /// element, attribute and piece of text writes, for every kind of text: entities, white space at either end, line breaks,
    /// characters beyond ASCII, an empty text, and texts around the size of a chunk.
    /// </summary>
    [Fact]
    public void TheSharedStringsPart_IsWrittenExactlyAsXmlWriterWritesIt()
    {
        var texts = new List<string> { "", " ", "plain", " lead", "trail ", "\tTabbed\t", "A & B", "<tag>", "a>b<c&d", "\"quoted\" 'single'", "line\r\nbreak\nand\rmore", "é€😀 ñ", "x&", "&x", "<", ">", "&" };
        foreach (var length in new[] { 1600, 1640, 1641, 1642, 1700, 8000, 8100, 20000 })
        {
            texts.Add(new string('a', length));
            texts.Add(new string('&', length / 6) + "tail");
            texts.Add("é" + new string('b', length) + "<");
        }
        for (var i = 0; i < 3000; i++) texts.Add("entry " + i + (i % 7 == 0 ? " & <" + i + ">" : "") + (i % 11 == 0 ? " " : ""));

        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            foreach (var text in texts) s.Row().WriteCell(text);
        }, new PackageOptions { StringStorage = StringStorage.Shared, CompressionLevel = CompressionLevel.Fastest });

        var output = new MemoryStream();
        using (var w = XmlWriter.Create(output, new XmlWriterSettings { Encoding = new UTF8Encoding(false), CloseOutput = false, NewLineHandling = NewLineHandling.Entitize })) // Entitize: a CR in text is written as &#xD;, which is what the table writes.
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var distinct = texts.Distinct(StringComparer.Ordinal).ToList();
            w.WriteAttributeString("count", texts.Count.ToString());
            w.WriteAttributeString("uniqueCount", distinct.Count.ToString());
            foreach (var text in distinct)
            {
                w.WriteStartElement("si");
                w.WriteStartElement("t");
                if (XmlTextUtil.NeedsSpacePreserve(text)) w.WriteAttributeString("xml", "space", null, "preserve");
                w.WriteString(text);
                w.WriteEndElement();
                w.WriteEndElement();
            }
            w.WriteEndElement();
            w.WriteEndDocument();
        }
        Assert.Equal(new UTF8Encoding(false).GetString(output.ToArray()), Part(bytes, "xl/sharedStrings.xml"));
    }

    // ---- the second look: columns that repeat among many thousands of values ----------------------------------------

    /// <summary>A fixed pool of <paramref name="count"/> random lower-case tokens of <paramref name="length"/> characters, each starting with its number so that all are different.</summary>
    private static string[] TokenPool(int count, int length, int seed)
    {
        var random = new Random(seed); // A fixed seed: never string.GetHashCode, which differs from one process to the next.
        var pool = new string[count];
        var chars = new char[length];
        for (var i = 0; i < count; i++)
        {
            for (var k = 0; k < length; k++) chars[k] = (char)('a' + random.Next(26));
            var number = i.ToString(System.Globalization.CultureInfo.InvariantCulture);
            number.CopyTo(0, chars, 0, number.Length);
            pool[i] = new string(chars);
        }
        return pool;
    }

    /// <summary>
    /// The case the first sample cannot see: a column draws from 20,000 different values, so none repeats in the first
    /// cells and every repeat is too far back for the zip compressor (its window is 32 KB, about a hundred rows). It must
    /// still end up shared, early enough that the file is within about 10% of what <see cref="StringStorage.Shared"/> makes.
    /// </summary>
    [Theory]
    [InlineData(25)]
    [InlineData(80)]
    public void Auto_AColumnThatRepeatsAmongTwentyThousandValues_IsSharedAfterASecondLook(int length)
    {
        const int rows = 100_000;
        var poolA = TokenPool(20_000, length, 101);
        var poolB = TokenPool(20_000, length, 202);
        string[] Row(int i) => new[] { poolA[Pick(i, 20_000)], poolB[Pick(rows + i, 20_000)] };

        byte[] Build(StringStorage storage) => ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++)
            {
                var values = Row(i);
                s.Row().WriteCell(values[0]).WriteCell(values[1]);
            }
        }, Options(storage));

        var auto = Build(StringStorage.Auto);
        var shared = Build(StringStorage.Shared);
        Assert.True(auto.Length <= shared.Length * 1.10, $"Auto made {auto.Length:N0} bytes, Shared {shared.Length:N0}.");

        // Both columns were looked at again and shared a few hundred cells in: only the cells before that are inline.
        var sheet = Part(auto, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"inlineStr\"") < 2 * (StringColumnPolicy.FirstLookAt + 2 * StringColumnPolicy.LookCells), $"{Count(sheet, "t=\"inlineStr\"")} inline cells.");
        AssertReadsBack(auto, Enumerable.Range(0, rows).Select(i => (string?[])Row(i)).ToList());
    }

    /// <summary>A column of text that never repeats stays inline after both looks, and what the looks put in the table is a few hundred texts, not the column.</summary>
    [Fact]
    public void Auto_AColumnOfUniqueText_StaysInlineAfterTheSecondLooks()
    {
        const int rows = 60_000;
        long uniqueInTable = 0;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell("Order SO-" + i.ToString("D8")).WriteCell("customer" + i + "@example.com");
            uniqueInTable = wb.SharedStrings.UniqueCount;
        }, Options(StringStorage.Auto));

        // Per column: the first cells up to the early exit, and one window of each of the two looks.
        Assert.True(uniqueInTable <= 2 * (StringColumnPolicy.EarlyExitCells + 2 * StringColumnPolicy.LookCells), $"The table holds {uniqueInTable} texts.");
        var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"inlineStr\"") >= 2 * rows - uniqueInTable);
        AssertReadsBack(bytes, Enumerable.Range(0, rows).Select(i => new[] { "Order SO-" + i.ToString("D8"), "customer" + i + "@example.com" }).ToList());
    }

    /// <summary>
    /// Values from a small set, however long the column, fit the compressor's window: a repeat is never far back, so nothing is
    /// gained from the table. A column of such values that starts as unique text and settles later stays inline.
    /// </summary>
    [Fact]
    public void Auto_AColumnOfAFewShortValues_ThatBeginsUnique_StaysInlineAfterTheLooks()
    {
        const int rows = 40_000;
        string Code(int i) => i < 200 ? "id" + i : "C" + Pick(i, 40).ToString("D2");
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell(Code(i)).WriteCell("ref" + i);
        }, Options(StringStorage.Auto));

        var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"inlineStr\"") > 2 * rows - 2 * (StringColumnPolicy.EarlyExitCells + 2 * StringColumnPolicy.LookCells + StringColumnPolicy.ProbeCells));
        AssertReadsBack(bytes, Enumerable.Range(0, rows).Select(i => new[] { Code(i), "ref" + i }).ToList());
    }

    /// <summary>A column that is unique for a while and then draws from many thousands of values is shared by the last look, in the middle of the sheet; every cell still reads back.</summary>
    [Fact]
    public void Auto_AColumnThatBeginsUniqueAndThenRepeatsAmongManyValues_ReadsBackExactlyAcrossTheChange()
    {
        const int rows = 60_000;
        var pool = TokenPool(12_000, 30, 303);
        string Text(int i) => i < 2000 ? "unique " + i : pool[Pick(i, 12_000)];
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 0; i < rows; i++) s.Row().WriteCell(Text(i)).WriteCell(i).WriteCell("after " + i);
        }, Options(StringStorage.Auto));

        var sheet = Part(bytes, "xl/worksheets/sheet1.xml");
        Assert.True(Count(sheet, "t=\"s\"") > rows / 2, "The last look shares the column from the middle of its first third on.");
        Assert.True(Count(sheet, "t=\"inlineStr\"") > rows, "The column of ids is inline.");
        var read = Read(bytes);
        Assert.Equal(rows, read.Count);
        for (var i = 0; i < rows; i += 7)
        {
            Assert.Equal(Text(i), read[i][0] as string);
            Assert.Equal("after " + i, read[i][2] as string);
        }
    }

    /// <summary>The same column through the two other ways of writing a sheet that reorder or buffer the cells: AutoFit (rows are spilled to a file first) and random access (<c>Cells[]</c>).</summary>
    [Fact]
    public void Auto_ASecondLook_ReadsBackExactly_WithAutoFitAndWithRandomAccessCells()
    {
        const int rows = 30_000;
        var pool = TokenPool(20_000, 40, 404);
        string Text(int i) => pool[Pick(i, 20_000)];

        var autoFit = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.AutoFitColumns();
            for (var i = 0; i < rows; i++) s.Row().WriteCell(Text(i)).WriteCell("row " + i);
        }, Options(StringStorage.Auto));
        Assert.True(Count(Part(autoFit, "xl/worksheets/sheet1.xml"), "t=\"s\"") > rows, "The column is shared.");
        AssertReadsBack(autoFit, Enumerable.Range(0, rows).Select(i => new[] { Text(i), "row " + i }).ToList());

        var random = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var i = 1; i <= rows; i++)
            {
                s.Cells[i, 2].Value = "row " + (i - 1);
                s.Cells[i, 1].Value = Text(i - 1);
            }
        }, Options(StringStorage.Auto));
        Assert.True(Count(Part(random, "xl/worksheets/sheet1.xml"), "t=\"s\"") > rows, "The column is shared.");
        AssertReadsBack(random, Enumerable.Range(0, rows).Select(i => new[] { Text(i), "row " + i }).ToList());
    }

    /// <summary>
    /// What the decision does with the kinds of column, without a workbook: whether a column becomes shared and how soon. The
    /// data is a fixed random draw from the given number of values of the given length, after a start of unique texts (so that the first judgement
    /// is "unique" and the looks decide).
    /// </summary>
    [Theory]
    [InlineData(0, 25, false)]       // unique text
    [InlineData(20_000, 25, true)]   // many values, far apart: a vocabulary of 500 KB
    [InlineData(20_000, 80, true)]
    [InlineData(100_000, 25, true)]  // more values still: caught by the last look
    [InlineData(1_500, 30, true)]    // a vocabulary of 45 KB
    [InlineData(1_000, 13, false)]   // 1,000 customer names: 13 KB, which the compressor remembers, and finds the words of again and again
    [InlineData(40, 7, false)]       // a few short values
    [InlineData(40, 30, false)]      // a few long values: 1 KB
    public void TheSecondLook_SharesWhatRepeatsBeyondTheCompressorWindow(int distinct, int length, bool sharedByTheLooks)
    {
        var random = new Random(77);
        var pool = distinct == 0 ? null : TokenPool(distinct, length, 5);
        string Text(int i) => pool == null || i < 128 ? "u" + i + new string('x', length) : pool[random.Next(distinct)];

        var table = new HashSet<string>();
        var state = new StringColumnState();
        var sharedAt = -1;
        for (var i = 0; i < 100_000; i++)
        {
            if (state.WritesInline()) continue;
            var text = Text(i);
            state.Observe(!table.Add(text), text.Length);
            if (sharedAt < 0 && state.Mode == ColumnMode.Shared) sharedAt = i;
        }
        Assert.Equal(sharedByTheLooks, sharedAt >= 0);
        if (distinct == 20_000) Assert.True(sharedAt <= StringColumnPolicy.FirstLookAt + StringColumnPolicy.LookCells + 8, $"Shared only at cell {sharedAt}.");
        if (!sharedByTheLooks) Assert.Equal(ColumnMode.Inline, state.Mode);
    }
    // ---- bounded memory ---------------------------------------------------------------------------------------------

    [Fact]
    public void Auto_TheStateGrowsWithTheColumnsAndNeverWithTheRows()
    {
        // Sixteen thousand columns of text in one row: the state is at most the 16,384 columns Excel has.
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            for (var r = 0; r < 3; r++)
            {
                var w = s.Row();
                for (var c = 0; c < 16384; c++) w.WriteCell("r" + r + "c" + c);
            }
        }, Options(StringStorage.Auto));
        var rows = Read(bytes);
        Assert.Equal(3, rows.Count);
        Assert.Equal("r2c16383", rows[2][16383]);
    }

    // ---- the decision itself ---------------------------------------------------------------------------------------

    [Fact]
    public void TheColumnState_SharesWhatRepeatsAndLetsGoOfWhatStopsRepeating()
    {
        // A repeating, long-text column is judged shared at the end of the first window and stays so while it keeps repeating.
        var repeating = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.ProbeCells - 1; i++) repeating.Observe(hit: i >= 8, length: 30);
        Assert.Equal(ColumnMode.Sampling, repeating.Mode);
        repeating.Observe(hit: true, length: 30);
        Assert.Equal(ColumnMode.Shared, repeating.Mode);
        for (var i = 0; i < 3 * StringColumnPolicy.MonitorCells; i++) repeating.Observe(hit: i % 3 != 0, length: 30);
        Assert.Equal(ColumnMode.Shared, repeating.Mode);

        // A unique column is judged unique at once, written inline without a lookup, and looked at again twice before it is inline for good.
        var unique = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.EarlyExitCells - 1; i++) unique.Observe(hit: false, length: 12);
        Assert.Equal(ColumnMode.Sampling, unique.Mode);
        unique.Observe(hit: false, length: 12);
        Assert.Equal(ColumnMode.Waiting, unique.Mode);
        var inlineCells = 0;
        while (unique.WritesInline()) inlineCells++;
        Assert.Equal(ColumnMode.Looking, unique.Mode);
        Assert.Equal(StringColumnPolicy.FirstLookAt - StringColumnPolicy.EarlyExitCells - 1, inlineCells);
        for (var i = 0; i < StringColumnPolicy.LookCells; i++) unique.Observe(hit: false, length: 12);
        Assert.Equal(ColumnMode.Waiting, unique.Mode);
        while (unique.WritesInline()) { }
        for (var i = 0; i < StringColumnPolicy.LookCells; i++) unique.Observe(hit: false, length: 12);
        Assert.Equal(ColumnMode.Inline, unique.Mode);
        Assert.True(unique.WritesInline());

        // A shared column whose repeats dry up is let go at the end of its monitoring window.
        var drying = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.ProbeCells; i++) drying.Observe(hit: i >= 8, length: 30);
        Assert.Equal(ColumnMode.Shared, drying.Mode);
        for (var i = 0; i < StringColumnPolicy.MonitorCells; i++) drying.Observe(hit: false, length: 30);
        Assert.Equal(ColumnMode.Inline, drying.Mode);

        // Short texts from a small set repeat but are not worth sharing; many distinct ones are.
        var few = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.ProbeCells; i++) few.Observe(hit: i >= 5, length: 6);
        Assert.Equal(ColumnMode.Inline, few.Mode);
        var many = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.ProbeCells; i++) many.Observe(hit: i % 2 == 1, length: 6); // Half of the 512 cells are new: 256 distinct texts.
        Assert.Equal(ColumnMode.Shared, many.Mode);

        // A handful of accidental repeats in a column of unique text is not enough to share it at the first judgement; it gets its second looks.
        var accidental = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.ProbeCells; i++) accidental.Observe(hit: i % 40 == 0, length: 20);
        Assert.Equal(ColumnMode.Waiting, accidental.Mode);
    }
    [Fact]
    public void TheConstants_AreTheOnesTheDocumentationNames()
    {
        // docs/index.md and the PackageOptions comments quote these numbers.
        Assert.Equal(128, StringColumnPolicy.EarlyExitCells);
        Assert.Equal(512, StringColumnPolicy.ProbeCells);
        Assert.Equal(5, StringColumnPolicy.EnterPercent);
        Assert.Equal(16, StringColumnPolicy.MinRepeatedLength);
        Assert.Equal(200, StringColumnPolicy.MinDistinct);
        Assert.Equal(512, StringColumnPolicy.FirstLookAt);
        Assert.Equal(8192, StringColumnPolicy.LastLookAt);
        Assert.Equal(512, StringColumnPolicy.LookCells);
        Assert.Equal(3, StringColumnPolicy.MinLookHits);
        Assert.Equal(24 * 1024, StringColumnPolicy.WindowBytes);
        Assert.Equal(100, StringColumnPolicy.MaxCellsPerRepeat);
    }
}
