using System.IO.Compression;
using System.Text;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// Reader support for the sheet metadata that follows the rows: merged ranges and hyperlinks (issue #6).
/// Workbooks are written with SheetStream and read back, so every assertion is an exact round trip.
/// </summary>
public class ReaderRangeTests
{
    private static byte[] LinksAndMerges() => ExcelPackage.CreateBytes(wb =>
    {
        var s = wb.Worksheets.Add("Report");
        s.MergeCells("A1:D1");
        s.MergeCells(3, 1, 5, 2);
        s.Row().WriteCell("Title spanning four columns");
        s.Row().WriteHyperlink("P4 Software", "https://p4software.com/", tooltip: "Open site")
               .WriteHyperlink("Mail", "mailto:info@example.com")
               .WriteInternalLink("Go to summary", "'Summary'!A1", tooltip: "Jump");
        s.AddHyperlink(10, 3, "https://example.com/x?y=1&z=2");
        s.AddInternalLink(11, 1, "Totals"); // A defined name as target.
        wb.Worksheets.Add("Summary").Row().WriteCell("target");
    });

    [Fact]
    public void MergedRanges_RoundTrip()
    {
        using var reader = ExcelReader.Open(new MemoryStream(LinksAndMerges()));
        var merges = reader.ReadMergedRanges("Report");

        Assert.Equal(new[] { "A1:D1", "A3:B5" }, merges.Select(m => m.Range));

        var title = merges[0];
        Assert.Equal((1, 1, 1, 4), (title.FromRow, title.FromColumn, title.ToRow, title.ToColumn));
        Assert.True(title.Contains(1, 1));
        Assert.True(title.Contains(1, 4));
        Assert.False(title.Contains(1, 5));
        Assert.False(title.Contains(2, 1));

        var block = merges[1];
        Assert.Equal((3, 1, 5, 2), (block.FromRow, block.FromColumn, block.ToRow, block.ToColumn));
        Assert.True(block.Contains(4, 2));
        Assert.False(block.Contains(6, 2));
        Assert.Equal("A3:B5", block.ToString());

        // The index overload and the name lookup (case-insensitive) see the same data.
        Assert.Equal(merges.Select(m => m.Range), reader.ReadMergedRanges(0).Select(m => m.Range));
        Assert.Equal(merges.Select(m => m.Range), reader.ReadMergedRanges("REPORT").Select(m => m.Range));
    }

    [Fact]
    public void Hyperlinks_RoundTrip()
    {
        using var reader = ExcelReader.Open(new MemoryStream(LinksAndMerges()));
        var links = reader.ReadHyperlinks("Report");

        Assert.Equal(5, links.Count);

        var site = links[0];
        Assert.Equal((2, 1, "A2"), (site.Row, site.Column, site.Address));
        Assert.Equal("https://p4software.com/", site.Url);
        Assert.Null(site.Location);
        Assert.Equal("Open site", site.Tooltip);
        Assert.Null(site.Display); // SheetStream does not write display text; the cell value is the visible text.

        var mail = links[1];
        Assert.Equal("B2", mail.Address);
        Assert.Equal("mailto:info@example.com", mail.Url);
        Assert.Null(mail.Tooltip);

        var jump = links[2];
        Assert.Equal("C2", jump.Address);
        Assert.Null(jump.Url);
        Assert.Equal("'Summary'!A1", jump.Location);
        Assert.Equal("Jump", jump.Tooltip);

        var late = links[3];
        Assert.Equal((10, 3, "C10"), (late.Row, late.Column, late.Address));
        Assert.Equal("https://example.com/x?y=1&z=2", late.Url); // The & survives XML escaping both ways.

        var name = links[4];
        Assert.Equal("A11", name.Address);
        Assert.Equal("Totals", name.Location);
        Assert.Equal("A11 -> Totals", name.ToString());

        Assert.Equal(links.Select(l => l.Address), reader.ReadHyperlinks(0).Select(l => l.Address));

        // The link cells themselves still read as ordinary text through the row API.
        var row2 = reader.ReadRows("Report").First(r => r.Index == 2);
        Assert.Equal("P4 Software", row2.GetString(0));
    }

    [Fact]
    public void SheetWithoutMergesOrLinks_ReturnsEmptyLists()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Plain");
            s.Row().WriteCells("a", 1.0, true);
            s.Row().WriteCells("b", 2.0, false);
            wb.Worksheets.Add("Empty"); // No rows at all: sheetData is an empty element.
        });

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        Assert.Empty(reader.ReadMergedRanges("Plain"));
        Assert.Empty(reader.ReadHyperlinks("Plain"));
        Assert.Empty(reader.ReadMergedRanges("Empty"));
        Assert.Empty(reader.ReadHyperlinks(1));
    }

    [Fact]
    public void UnknownSheet_ThrowsLikeReadRows()
    {
        using var reader = ExcelReader.Open(new MemoryStream(LinksAndMerges()));
        Assert.Throws<ArgumentException>(() => reader.ReadMergedRanges("Nope"));
        Assert.Throws<ArgumentException>(() => reader.ReadHyperlinks("Nope"));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.ReadMergedRanges(7));
        Assert.Throws<ArgumentOutOfRangeException>(() => reader.ReadHyperlinks(-1));
    }

    [Fact]
    public void LongSheet_MergeAndLinkAtTheEndAreFound()
    {
        const int rows = 20_000;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Long");
            for (var i = 1; i <= rows; i++)
                s.Row().WriteCell("Line " + i).WriteCell(i).WriteCell(i * 0.5);
            s.MergeCells(rows, 1, rows, 3);
            s.AddHyperlink(rows, 2, "https://example.com/last");
        });

        using var reader = ExcelReader.Open(new MemoryStream(bytes));

        var merge = Assert.Single(reader.ReadMergedRanges("Long"));
        Assert.Equal("A20000:C20000", merge.Range);
        Assert.True(merge.Contains(rows, 2));

        var link = Assert.Single(reader.ReadHyperlinks("Long"));
        Assert.Equal("B20000", link.Address);
        Assert.Equal("https://example.com/last", link.Url);
    }

    [Fact]
    public void MalformedEntries_AreIgnored_AndExcelOnlyForms_AreRead()
    {
        // Start from a valid file, then edit the sheet XML the way a damaged or Excel-authored file might look.
        var bytes = MutatePart(LinksAndMerges(), "xl/worksheets/sheet1.xml", xml => xml
            // Broken merges: not a reference, whole-row form, empty ref, and no ref at all.
            .Replace("<mergeCells count=\"2\">", "<mergeCells count=\"6\"><mergeCell ref=\"nonsense\" /><mergeCell ref=\"1:1\" /><mergeCell ref=\"\" /><mergeCell />")
            // Excel may write a ranged hyperlink ref and a display attribute; a dangling r:id must not throw.
            .Replace("ref=\"C10\"", "ref=\"C10:D11\" display=\"Late_x000A_link\"")
            .Replace("ref=\"A11\"", "ref=\"A11\" r:id=\"rId999\"")
            // A hyperlink with an unparseable ref is dropped.
            .Replace("</hyperlinks>", "<hyperlink ref=\"ZZZZZZ1\" location=\"x\" /></hyperlinks>"));

        using var reader = ExcelReader.Open(new MemoryStream(bytes));

        Assert.Equal(new[] { "A1:D1", "A3:B5" }, reader.ReadMergedRanges("Report").Select(m => m.Range));

        var links = reader.ReadHyperlinks("Report");
        Assert.Equal(5, links.Count);
        var late = links[3];
        Assert.Equal("C10", late.Address); // Top-left cell of the ranged ref.
        Assert.Equal("Late\nlink", late.Display); // _x000A_ decoded like cell text.
        Assert.Equal("https://example.com/x?y=1&z=2", late.Url);
        var name = links[4];
        Assert.Null(name.Url); // rId999 has no relationship: no URL, but the location survives.
        Assert.Equal("Totals", name.Location);
    }

    /// <summary>Rebuilds the package with one XML part rewritten by <paramref name="mutate"/>.</summary>
    private static byte[] MutatePart(byte[] package, string part, Func<string, string> mutate)
    {
        var ms = new MemoryStream();
        ms.Write(package, 0, package.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(part)!;
            string text;
            using (var sr = new StreamReader(entry.Open())) text = sr.ReadToEnd();
            entry.Delete();
            using var w = new StreamWriter(zip.CreateEntry(part).Open(), new UTF8Encoding(false));
            w.Write(mutate(text));
        }
        return ms.ToArray();
    }
}
