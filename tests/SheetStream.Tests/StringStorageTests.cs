using System.IO.Compression;
using System.Text;
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
            _picks = Enumerable.Range(0, Math.Max(i + 1, 100_000)).Select(_ => random.Next()).ToArray();
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
        Assert.True(uniqueInTable < rows, "The table did not take every text.");
        AssertReadsBack(bytes, expected);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        Assert.Empty(new OpenXmlValidator().Validate(doc));
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

        // A unique column is judged inline for good.
        var unique = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.EarlyExitCells - 1; i++) unique.Observe(hit: false, length: 12);
        Assert.Equal(ColumnMode.Sampling, unique.Mode);
        unique.Observe(hit: false, length: 12);
        Assert.Equal(ColumnMode.Inline, unique.Mode);

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

        // A handful of accidental repeats in a column of unique text is not enough.
        var accidental = new StringColumnState();
        for (var i = 0; i < StringColumnPolicy.ProbeCells; i++) accidental.Observe(hit: i % 40 == 0, length: 20);
        Assert.Equal(ColumnMode.Inline, accidental.Mode);
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
    }
}
