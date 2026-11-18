using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using SheetStream.Reading;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Tests;

/// <summary>
/// <see cref="PackageOptions.WriteCellReferences"/>: off (the default) a cell in the column after the previous one
/// has no <c>r</c> attribute; on, every cell has one, in every way of writing cells. Both settings must read back
/// to identical values.
/// </summary>
public class WriteCellReferencesTests
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string SheetPart = "xl/worksheets/sheet1.xml";

    private static PackageOptions On => new() { WriteCellReferences = true };

    /// <summary>
    /// A dense row, a Skip, a leading Skip, an unstyled empty cell and a styled empty cell: every shape in which
    /// a cell is, or is not, the next one in line.
    /// </summary>
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

    /// <summary>The <c>r</c> attribute of every cell of the first sheet, row by row (null where it is omitted).</summary>
    private static List<string?[]> References(byte[] workbook)
    {
        using var zip = new ZipArchive(new MemoryStream(workbook), ZipArchiveMode.Read);
        using var stream = zip.GetEntry(SheetPart)!.Open();
        return XDocument.Load(stream).Descendants(Main + "row")
            .Select(row => row.Elements(Main + "c").Select(c => (string?)c.Attribute("r")).ToArray())
            .ToList();
    }

    private static List<object?[]> ReadAll(byte[] workbook)
    {
        using var reader = ExcelReader.Open(new MemoryStream(workbook));
        return reader.ReadRows("S").Select(r => r.Values.ToArray()).ToList();
    }

    private static void AssertEveryCellHasAReference(byte[] workbook)
    {
        var rows = References(workbook);
        Assert.NotEmpty(rows);
        Assert.All(rows.SelectMany(r => r), r => Assert.NotNull(r));
    }

    private static byte[] Edit(byte[] input, Action<Worksheet> change, PackageOptions? options)
    {
        using var output = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(input), output, leaveOpen: true, options))
        {
            change(package.Workbook.Worksheets.Edit("S"));
            package.Save();
        }
        return output.ToArray();
    }

    [Fact]
    public void TheDefaultIsOff()
    {
        Assert.False(new PackageOptions().WriteCellReferences);
    }

    [Fact]
    public void Off_ConsecutiveCellsHaveNoReference_AndGapsKeepIt()
    {
        var bytes = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));
        var refs = References(bytes);

        Assert.Equal(new string?[] { null, null, null }, refs[0]);
        Assert.Equal(new string?[] { null, "D2" }, refs[1]);
        Assert.Equal(new string?[] { "B3", "D3" }, refs[2]);
        Assert.Equal(new string?[] { null, null, null }, refs[3]);
    }

    [Fact]
    public void On_Streaming_EveryCellHasItsAddress()
    {
        var bytes = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")), On);
        var refs = References(bytes);

        Assert.Equal(new string?[] { "A1", "B1", "C1" }, refs[0]);
        Assert.Equal(new string?[] { "A2", "D2" }, refs[1]);
        Assert.Equal(new string?[] { "B3", "D3" }, refs[2]);          // The unstyled empty cell still writes no cell at all.
        Assert.Equal(new string?[] { "A4", "B4", "C4" }, refs[3]);
    }

    [Fact]
    public void On_EveryKindOfCell_HasItsAddress()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell("text").WriteCell(1.5).WriteCell(2).WriteCell(true)
                .WriteCell(new DateTime(2026, 1, 2)).WriteFormula("A1&\"x\"").WriteCell(TimeSpan.FromHours(3));
        }, On);

        Assert.Equal(new string?[] { "A1", "B1", "C1", "D1", "E1", "F1", "G1" }, References(bytes).Single());
    }

    [Fact]
    public void On_WithInlineStrings_EveryCellHasItsAddress()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell("a").WriteCell("b").WriteCell("c");
        }, new PackageOptions { WriteCellReferences = true, UseSharedStrings = false });

        Assert.Equal(new string?[] { "A1", "B1", "C1" }, References(bytes).Single());
    }

    [Fact]
    public void On_AutoFit_EveryCellHasItsAddress()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.AutoFitColumns();
            FillStreaming(wb, s);
        }, On);

        AssertEveryCellHasAReference(bytes);
        Assert.Equal(new string?[] { "A2", "D2" }, References(bytes)[1]);
    }

    [Fact]
    public void On_RandomAccessCells_EveryCellHasItsAddress()
    {
        var bytes = ExcelPackage.CreateBytes(wb => FillByCells(wb, wb.Worksheets.Add("S")), On);

        var refs = References(bytes);
        Assert.Equal(new string?[] { "A1", "B1", "C1" }, refs[0]);
        Assert.Equal(new string?[] { "A2", "D2" }, refs[1]);
        Assert.Equal(new string?[] { "B3", "D3" }, refs[2]);
        Assert.Equal(new string?[] { "A4", "B4", "C4" }, refs[3]);
    }

    [Fact]
    public void On_Edit_EveryCellHasItsAddress_EvenWhenTheSourceHadNone()
    {
        // The source was written with the default, so its cells have no references; the edited copy has them all.
        var original = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));
        Assert.Contains(References(original).SelectMany(r => r), r => r == null);

        var edited = Edit(original, s =>
        {
            s.Cells["C2"].Value = "filled";
            s.Cells["D4"].Value = "tail";
        }, On);

        AssertEveryCellHasAReference(edited);
        var rows = ReadAll(edited);
        Assert.Equal(new object?[] { "x", null, "filled", "y" }, rows[1]);
        Assert.Equal(new object?[] { null, 2.5, 3.5, "tail" }, rows[3]);
    }

    [Fact]
    public void Off_Edit_StillLeavesTheAddressOffConsecutiveCells()
    {
        var original = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));
        var edited = Edit(original, s => s.Cells["A1"].Value = "changed", options: null);

        Assert.Equal(new string?[] { null, null, null }, References(edited)[0]);
    }

    [Fact]
    public void On_Template_RowsFilledIntoItHaveTheirAddress()
    {
        // Built with the option on too: rows a template contributes are copied unchanged, so only then is every
        // cell of the filled sheet expected to carry an address.
        var template = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell("Report");
            s.Row(3).WriteCell("Sku").WriteCell("Qty");
        }, On);

        using var output = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true, On))
        {
            var s = package.Workbook.Worksheets.Open("S");
            s.Row(4).WriteCell("A-1").WriteCell(5).WriteCell(true);
            s.Row(5).WriteCell("B-2").Skip().WriteCell(7);
            package.Save();
        }

        var refs = References(output.ToArray());
        Assert.Equal(new string?[] { "A4", "B4", "C4" }, refs[2]);
        Assert.Equal(new string?[] { "A5", "C5" }, refs[3]);
        AssertEveryCellHasAReference(output.ToArray());
    }

    [Fact]
    public void Off_Template_RowsFilledIntoItHaveNoAddressOnConsecutiveCells()
    {
        var template = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S").Row(1).WriteCell("Report"));

        using var output = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true))
        {
            package.Workbook.Worksheets.Open("S").Row(2).WriteCell("A-1").WriteCell(5);
            package.Save();
        }

        Assert.Equal(new string?[] { null, null }, References(output.ToArray())[1]);
    }

    [Fact]
    public void BothSettings_ReadBackIdentically()
    {
        void Build(Workbook wb) => FillStreaming(wb, wb.Worksheets.Add("S"));
        var off = ExcelPackage.CreateBytes(Build);
        var on = ExcelPackage.CreateBytes(Build, On);

        // Guard the test itself: the two files really differ in whether the address is there.
        Assert.Contains(References(off).SelectMany(r => r), r => r == null);
        AssertEveryCellHasAReference(on);

        var expected = ReadAll(off);
        var readOn = ReadAll(on);
        Assert.Equal(expected.Count, readOn.Count);
        for (var i = 0; i < expected.Count; i++) Assert.Equal(expected[i], readOn[i]);

        // The random-access path too.
        var byCellsOff = ReadAll(ExcelPackage.CreateBytes(wb => FillByCells(wb, wb.Worksheets.Add("S"))));
        var byCellsOn = ReadAll(ExcelPackage.CreateBytes(wb => FillByCells(wb, wb.Worksheets.Add("S")), On));
        for (var i = 0; i < expected.Count; i++)
        {
            Assert.Equal(expected[i], byCellsOff[i]);
            Assert.Equal(expected[i], byCellsOn[i]);
        }
    }

    [Fact]
    public void On_OpenXmlSdk_CellReference_IsNeverNull()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            FillStreaming(wb, wb.Worksheets.Add("S"));
            var wide = wb.Worksheets.Add("Wide");
            wide.Row().WriteCell("a").WriteCell(1).WriteCell(2).WriteCell("d").WriteCell(false);
        }, On);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var cells = doc.WorkbookPart!.WorksheetParts
            .SelectMany(p => p.Worksheet!.Descendants<X.Cell>()).ToList();

        Assert.NotEmpty(cells);
        Assert.All(cells, c => Assert.NotNull(c.CellReference?.Value));
    }

    [Fact]
    public void Off_OpenXmlSdk_CellReference_IsNullForConsecutiveCells()
    {
        // Pins the trade-off the option exists for: if this ever stops being true the option's documentation is stale.
        var bytes = ExcelPackage.CreateBytes(wb => FillStreaming(wb, wb.Worksheets.Add("S")));

        using var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var cells = doc.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<X.Cell>().ToList();

        Assert.Contains(cells, c => c.CellReference?.Value == null);
    }
}
