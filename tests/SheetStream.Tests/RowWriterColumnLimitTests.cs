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
