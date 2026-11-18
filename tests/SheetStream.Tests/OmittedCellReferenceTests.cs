using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Internal;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// A cell that sits in the column after the previous one is written without its <c>r</c> attribute (ECMA-376
/// makes it optional). These tests pin the other half of that rule: whenever a cell is NOT the next one in line
/// the reference is still written, and every way of reading the file (the reader, edit mode, template mode)
/// puts every value back in exactly the column it was written to.
/// </summary>
public class OmittedCellReferenceTests
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string SheetPart = "xl/worksheets/sheet1.xml";

    /// <summary>The values the sheet written by <see cref="FillStreaming"/> must read back as, row by row.</summary>
    private static readonly object?[][] Expected =
    {
        new object?[] { "a", 1.0, true },                       // Three consecutive cells: no references at all.
        new object?[] { "x", null, null, "y" },                 // Skip(2): the cell after the gap needs its reference.
        new object?[] { null, 5.0, null, 7.0 },                 // Leading Skip, then an unstyled empty cell, then a value.
        new object?[] { null, 2.5, 3.5 },                       // Styled empty first cell is written (as a cell without a value), reads as empty.
    };

    /// <summary>Row-by-row writes covering a dense row, Skip, an unstyled empty cell and a styled empty cell.</summary>
    private static void FillStreaming(Workbook wb, Worksheet s)
    {
        var fill = wb.Styles.Create(o => o.Fill.BackgroundColor = "#DDEBF7");
        s.Row(1).WriteCell("a").WriteCell(1).WriteCell(true);
        s.Row(2).WriteCell("x").Skip(2).WriteCell("y");
        s.Row(3).Skip().WriteCell(5).WriteEmpty().WriteCell(7);
        s.Row(4).WriteEmpty(fill).WriteCell(2.5).WriteCell(3.5);
    }

    /// <summary>The same content as <see cref="FillStreaming"/>, written by cell address in a scrambled order.</summary>
    private static void FillByCells(Workbook wb, Worksheet s)
    {
        var fill = wb.Styles.Create(o => o.Fill.BackgroundColor = "#DDEBF7");
        s.Cells[4, 3].Value = 3.5;
        s.Cells[3, 4].Value = 7;
        s.Cells[1, 1].Value = "a";
        s.Cells[2, 4].Value = "y";
        s.Cells[4, 1].Style = fill;
        s.Cells[3, 2].Value = 5;
        s.Cells[1, 3].Value = true;
        s.Cells[4, 2].Value = 2.5;
        s.Cells[2, 1].Value = "x";
        s.Cells[1, 2].Value = 1;
    }

    private static XDocument SheetXml(byte[] workbook)
    {
        using var zip = new ZipArchive(new MemoryStream(workbook), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(SheetPart)!.Open();
        return XDocument.Load(stream);
    }

    /// <summary>The <c>r</c> attribute of every cell of every row (null when omitted), by row number.</summary>
    private static Dictionary<int, List<string?>> ReferencesByRow(byte[] workbook)
        => SheetXml(workbook).Descendants(Main + "row").ToDictionary(
            row => (int)row.Attribute("r")!,
            row => row.Elements(Main + "c").Select(c => (string?)c.Attribute("r")).ToList());

    private static List<object?[]> ReadAll(byte[] workbook, string sheet = "S")
    {
        using var reader = ExcelReader.Open(new MemoryStream(workbook));
        return reader.ReadRows(sheet).Select(r => r.Values.ToArray()).ToList();
    }

    /// <summary>Rewrites every cell of the sheet to carry an explicit reference, the way the library used to write it.</summary>
    private static byte[] WithExplicitReferences(byte[] workbook)
    {
        var ms = new MemoryStream();
        ms.Write(workbook, 0, workbook.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var doc = SheetXml(workbook);
            foreach (var row in doc.Descendants(Main + "row"))
            {
                var column = 0;
                foreach (var cell in row.Elements(Main + "c"))
                {
                    var reference = (string?)cell.Attribute("r");
                    column = reference != null ? RangeReference.ParseCell(reference).Column : column + 1;
                    cell.SetAttributeValue("r", CellReference.ToA1((int)row.Attribute("r")!, column));
                }
            }
            zip.GetEntry(SheetPart)!.Delete();
            using var w = zip.CreateEntry(SheetPart).Open();
            doc.Save(w);
        }
        return ms.ToArray();
    }

    private static void AssertRowsEqual(IReadOnlyList<object?[]> expected, IReadOnlyList<object?[]> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        for (var i = 0; i < expected.Count; i++) Assert.Equal(expected[i], actual[i]);
    }

    [Fact]
    public void ConsecutiveCells_HaveNoReference_AndEveryGapKeepsIt()
    {
        var bytes = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));
        var refs = ReferencesByRow(bytes);

        Assert.Equal(new string?[] { null, null, null }, refs[1]);
        Assert.Equal(new string?[] { null, "D2" }, refs[2]);           // x at A, then the gap: y is D2.
        Assert.Equal(new string?[] { "B3", "D3" }, refs[3]);           // Leading Skip; the unstyled empty cell writes nothing, so D3 needs it too.
        Assert.Equal(new string?[] { null, null, null }, refs[4]);     // A styled empty cell is a real cell: the row stays consecutive.
    }

    [Fact]
    public void Streaming_ReadsBackInTheColumnsItWasWrittenTo()
    {
        var bytes = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));
        AssertRowsEqual(Expected, ReadAll(bytes));
    }

    [Fact]
    public void ReadingWithoutReferences_IsIdenticalToReadingTheSameSheetWithThem()
    {
        var bytes = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));
        var explicitRefs = WithExplicitReferences(bytes);

        // Guard the test itself: the rewrite really did add references where there were none.
        Assert.All(ReferencesByRow(explicitRefs).Values.SelectMany(r => r), r => Assert.NotNull(r));
        AssertRowsEqual(ReadAll(explicitRefs), ReadAll(bytes));
    }

    [Fact]
    public void ReaderHandlesMixedReferencedAndUnreferencedCellsInOneRow()
    {
        // A file from another writer may mix both: a cell with a reference moves the position to itself, and an
        // unreferenced cell after it is the next column along.
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell(1).WriteCell(2).WriteCell(3).WriteCell(4);
        });
        var patched = Patch(bytes, xml => xml.Replace("<c><v>3</v></c>", "<c r=\"F1\"><v>3</v></c>"));
        // A1, B1, F1, then the unreferenced fourth cell follows F1 as G1.
        Assert.Equal(new object?[] { 1.0, 2.0, null, null, null, 3.0, 4.0 }, ReadAll(patched).Single());
    }

    [Fact]
    public void AutoFit_WritesTheSameCellsAndReadsBackTheSame()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.AutoFitColumns();
            FillStreaming(wb, s);
        });

        Assert.Equal(new string?[] { null, "D2" }, ReferencesByRow(bytes)[2]);   // Rows went through the spill file and came back intact.
        AssertRowsEqual(Expected, ReadAll(bytes));
    }

    [Fact]
    public void RandomAccessCells_ReadBackAndOnlyGapsCarryReferences()
    {
        var bytes = ExcelPackage.CreateBytes(wb => FillByCells(wb, wb.Worksheets.Add("S")));

        var refs = ReferencesByRow(bytes);
        Assert.Equal(new string?[] { null, null, null }, refs[1]);
        Assert.Equal(new string?[] { null, "D2" }, refs[2]);
        Assert.Equal(new string?[] { "B3", "D3" }, refs[3]);
        Assert.Equal(new string?[] { null, null, null }, refs[4]);
        AssertRowsEqual(Expected, ReadAll(bytes));
    }

    [Fact]
    public void Edit_KeepsEveryValueInItsColumn_AndPlacesNewCellsExactly()
    {
        var original = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));

        // Opening and saving without a change must not move anything.
        AssertRowsEqual(Expected, ReadAll(Edit(original, _ => { })));

        var edited = Edit(original, s =>
        {
            Assert.Equal(7.0, s.Cells["D3"].Value);       // Found by address although the file has no reference to compare.
            Assert.Equal("y", s.Cells[2, 4].Value);
            Assert.Equal(true, s.Cells["C1"].Value);
            Assert.Null(s.Cells["C2"].Value);
            s.Cells["C2"].Value = "filled";               // Fills a gap between two cells.
            s.Cells["F2"].Value = 99;                     // Opens a new gap after the last cell.
            s.Cells["A3"].Value = "lead";                 // New first cell in a row that began at B.
            s.Cells["D4"].Value = "tail";                 // Extends a dense row.
        });

        var rows = ReadAll(edited);
        Assert.Equal(new object?[] { "a", 1.0, true }, rows[0]);
        Assert.Equal(new object?[] { "x", null, "filled", "y", null, 99.0 }, rows[1]);
        Assert.Equal(new object?[] { "lead", 5.0, null, 7.0 }, rows[2]);
        Assert.Equal(new object?[] { null, 2.5, 3.5, "tail" }, rows[3]);

        // And what the editor wrote is itself valid.
        using var doc = SpreadsheetDocument.Open(new MemoryStream(edited), false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc));
    }

    [Fact]
    public void Edit_OfASheetWithExplicitReferences_StillWorks()
    {
        // The other side of compatibility: files written before this change (and by Excel) have references everywhere.
        var original = WithExplicitReferences(ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S"))));
        var edited = Edit(original, s => s.Cells["B2"].Value = "new");
        var rows = ReadAll(edited);
        Assert.Equal(new object?[] { "x", "new", null, "y" }, rows[1]);
        Assert.Equal(Expected[0], rows[0]);
        Assert.Equal(Expected[2], rows[2]);
    }

    [Fact]
    public void TemplateStyle_FindsCellsThatHaveNoReference()
    {
        var template = ExcelPackage.CreateBytes(wb =>
        {
            var a = wb.Styles.Create(o => o.Font.Bold = true);
            var b = wb.Styles.Create(o => o.NumberFormat = "0.00");
            var c = wb.Styles.Create(o => o.Font.Italic = true);
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell("head");
            s.Row(2).WriteCell("a", a).WriteCell(1, b);              // A2 and B2 both without a reference.
            s.Row(3).Skip().WriteCell(2, c);                         // B3 with one.
            s.Row(4).WriteCell(1).WriteCell(2).WriteCell(3, c);      // C4 without one, after cells without.
        });
        int StyleOf(string cell, byte[] file)
        {
            using var ms = new MemoryStream();
            using var package = ExcelPackage.OpenTemplate(new MemoryStream(file), ms, leaveOpen: true);
            return package.Workbook.Worksheets.Open("S").TemplateStyle(cell).Id;
        }

        var a2 = StyleOf("A2", template);
        var b2 = StyleOf("B2", template);
        var b3 = StyleOf("B3", template);
        var c4 = StyleOf("C4", template);
        Assert.True(a2 > 0 && b2 > 0 && b3 > 0 && c4 > 0);
        Assert.NotEqual(a2, b2);
        Assert.NotEqual(b2, b3);
        Assert.Equal(b3, c4);                       // Same style object, written in two different ways.
        Assert.Equal(0, StyleOf("A3", template));    // Gap before B3.
        Assert.Equal(0, StyleOf("A4", template));    // Unstyled cell without a reference.
        Assert.Equal(0, StyleOf("D4", template));    // Past the last cell.
    }

    [Fact]
    public void Hyperlinks_UseTheRightCellAlthoughTheCellHasNoReference()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell("left").WriteHyperlink("docs", "https://example.com/1");
        });
        var link = SheetXml(bytes).Descendants(Main + "hyperlink").Single();
        Assert.Equal("B1", (string?)link.Attribute("ref"));
    }

    private static byte[] Edit(byte[] input, Action<Worksheet> change)
    {
        using var output = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(input), output, leaveOpen: true))
        {
            change(package.Workbook.Worksheets.Edit("S"));
            package.Save();
        }
        return output.ToArray();
    }

    private static byte[] Patch(byte[] workbook, Func<string, string> patch)
    {
        var ms = new MemoryStream();
        ms.Write(workbook, 0, workbook.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            string xml;
            using (var sr = new StreamReader(zip.GetEntry(SheetPart)!.Open())) xml = sr.ReadToEnd();
            zip.GetEntry(SheetPart)!.Delete();
            using var w = new StreamWriter(zip.CreateEntry(SheetPart).Open());
            w.Write(patch(xml));
        }
        return ms.ToArray();
    }
}
