using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Internal;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// <c>RowWriter.Skip</c> moves the cursor, and a cursor past Excel's last column (XFD, 16,384) must never reach the
/// bytes of a cell. A call that is refused has to leave the row exactly as it was: a caller who catches the exception
/// and goes on writing must still get a sheet Excel opens.
/// </summary>
public class RowWriterColumnLimitTests
{
    public static IEnumerable<object[]> Modes()
    {
        foreach (var storage in new[] { StringStorage.Auto, StringStorage.Shared, StringStorage.Inline })
            foreach (var references in new[] { false, true })
                yield return new object[] { storage, references };
    }

    private static PackageOptions Options(StringStorage storage, bool references) => new PackageOptions
    {
        StringStorage = storage,
        WriteCellReferences = references,
    };

    /// <summary>Opens the workbook with the Open XML SDK and fails on any schema error, then returns every sheet row as read by <see cref="ExcelReader"/>.</summary>
    private static List<object?[]> ValidateAndRead(byte[] workbook)
    {
        using (var doc = SpreadsheetDocument.Open(new MemoryStream(workbook), false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        }
        using var reader = ExcelReader.Open(new MemoryStream(workbook), false);
        return reader.ReadRows("S").Select(r => r.Values.ToArray()).ToList();
    }

    /// <summary>The repro of the review: an enormous Skip, the exception swallowed, then more writes and a Save.</summary>
    [Theory]
    [MemberData(nameof(Modes))]
    public void SkipThatWouldWrapTheCursor_IsRefused_AndTheRowIsUntouched(StringStorage storage, bool references)
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            var r = s.Row();
            r.WriteCell("a");
            try { r.Skip(int.MaxValue); r.WriteCell(42); } catch (ArgumentOutOfRangeException) { }
            Assert.Equal(2, r.NextColumn); // The refused call moved nothing.
            try { r.WriteCell("b"); r.Skip(int.MaxValue - 1); r.WriteCell("c"); } catch (ArgumentOutOfRangeException) { }
            r.WriteCell("d");
            for (var i = 0; i < 700; i++) s.Row().WriteCell("x" + i).WriteCell(i).WriteCell("y" + i % 3);
        }, Options(storage, references));

        var rows = ValidateAndRead(bytes);
        Assert.Equal(new object?[] { "a", "b", "d" }, rows[0]);
        Assert.Equal(701, rows.Count);
        Assert.Equal(new object?[] { "x699", 699.0, "y0" }, rows[700]);
    }

    /// <summary>Skip may bring the cursor to the column after XFD (nothing can be written there, but the row is still complete), and no further.</summary>
    [Theory]
    [MemberData(nameof(Modes))]
    public void SkipNearTheLastColumn_StopsAtTheEdge(StringStorage storage, bool references)
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");

            // Row 1: the last column is reachable and takes a value.
            var r1 = s.Row().WriteCell("first");
            r1.Skip(CellReference.MaxColumns - 2);
            Assert.Equal(CellReference.MaxColumns, r1.NextColumn);
            r1.WriteCell("last");
            Assert.Equal(CellReference.MaxColumns + 1, r1.NextColumn);

            // Row 2: one column too far is refused, with the cursor where it was.
            var r2 = s.Row().WriteCell("first");
            Assert.Throws<ArgumentOutOfRangeException>(() => r2.Skip(CellReference.MaxColumns));
            Assert.Equal(2, r2.NextColumn);
            r2.Skip(CellReference.MaxColumns - 1); // Exactly the column after XFD.
            Assert.Equal(CellReference.MaxColumns + 1, r2.NextColumn);
            Assert.Throws<ArgumentOutOfRangeException>(() => r2.Skip());

            // Every kind of cell is refused there, before a byte of it is staged.
            Assert.Throws<InvalidOperationException>(() => r2.WriteCell("text"));
            Assert.Throws<InvalidOperationException>(() => r2.WriteInlineString("text"));
            Assert.Throws<InvalidOperationException>(() => r2.WriteCell(1.5));
            Assert.Throws<InvalidOperationException>(() => r2.WriteCell(7));
            Assert.Throws<InvalidOperationException>(() => r2.WriteCell(true));
            Assert.Throws<InvalidOperationException>(() => r2.WriteFormula("1+1"));
            Assert.Throws<InvalidOperationException>(() => r2.WriteHyperlink("x", "https://example.com"));
            Assert.Throws<InvalidOperationException>(() => r2.WriteRichText(new RichText("x")));
            Assert.Equal(CellReference.MaxColumns + 1, r2.NextColumn);

            s.Row().WriteCell("after");
        }, Options(storage, references));

        var rows = ValidateAndRead(bytes);
        Assert.Equal(3, rows.Count);
        Assert.Equal("first", rows[0][0]);
        Assert.Equal("last", rows[0][CellReference.MaxColumns - 1]);
        Assert.Equal(new object?[] { "first" }, rows[1]);
        Assert.Equal(new object?[] { "after" }, rows[2]);
    }

    /// <summary>
    /// The runs of a rich-text cell are rendered into a buffer that the next rich-text cell reuses. A refused cell (here: no column left)
    /// must not leave its runs in it, or they come out in front of the next cell's text.
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RefusedRichText_LeavesNoRunsForTheNextRichTextCell(bool template)
    {
        byte[] Write(Func<Worksheet> sheet, Action<Worksheet> end)
        {
            var s = sheet();
            var r = s.Row(1).Skip(CellReference.MaxColumns - 1);
            r.WriteCell("last"); // Column XFD: the next cell has no column left.
            Assert.Throws<InvalidOperationException>(() => r.WriteRichText(new RichText("STALE") { Bold = true }));
            Assert.Throws<ArgumentException>(() => s.Row(2).WriteRichText(new RichText(new string('x', 32768))));
            s.Row(3).WriteRichText(new RichText("fresh"));
            end(s);
            return Array.Empty<byte>();
        }

        byte[] bytes;
        if (!template)
            bytes = ExcelPackage.CreateBytes(wb => Write(() => wb.Worksheets.Add("S"), _ => { }));
        else
        {
            var source = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S"));
            var output = new MemoryStream();
            using (var package = ExcelPackage.OpenTemplate(new MemoryStream(source), output, leaveOpen: true))
            {
                Write(() => package.Workbook.Worksheets.Open("S"), _ => { });
                package.Save();
            }
            bytes = output.ToArray();
        }

        var rows = ValidateAndRead(bytes);
        Assert.Equal("fresh", rows.Last()[0]);
    }

    /// <summary>Empty cells that write nothing must not carry the cursor past the column after XFD, which <c>Skip</c> promises it never passes.</summary>
    [Fact]
    public void UnstyledEmptyCells_StopTheCursorAtTheColumnAfterTheLast()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            var r = s.Row().Skip(CellReference.MaxColumns - 1).WriteCell("last");
            Assert.Equal(CellReference.MaxColumns + 1, r.NextColumn);
            r.WriteEmpty().WriteEmpty().WriteCell((string?)null).WriteCell((object?)null);
            Assert.Equal(CellReference.MaxColumns + 1, r.NextColumn);
            r.Skip(0); // Allowed where the cursor is.
            Assert.Throws<ArgumentOutOfRangeException>(() => r.Skip());
            s.Row().WriteCell("next");
        });
        var rows = ValidateAndRead(bytes);
        Assert.Equal(2, rows.Count);
        Assert.Equal("last", rows[0][CellReference.MaxColumns - 1]);
    }

    /// <summary>A hyperlink registered for a column that cannot hold it must not be left behind, or the sheet's trailer would carry a reference past XFD.</summary>
    [Fact]
    public void RefusedHyperlink_LeavesNoLinkBehind()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            var r = s.Row().Skip(CellReference.MaxColumns);
            Assert.Throws<InvalidOperationException>(() => r.WriteHyperlink("x", "https://example.com"));
            Assert.Throws<InvalidOperationException>(() => r.WriteInternalLink("x", "S!A1"));
            s.Row().WriteCell("ok");
        });

        var rows = ValidateAndRead(bytes);
        Assert.Equal(2, rows.Count);
    }
}
