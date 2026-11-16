using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// Excel refuses to open a workbook that holds a cell of more than 32,767 characters ("Unable to get the Open
/// property"), and a file written with such a cell used to save without a word. Every way of putting text in a cell
/// now refuses it, in plain words, before anything of the cell is written, and the writer stays usable.
/// </summary>
public class CellTextLengthTests
{
    private const int Max = 32_767;

    public static IEnumerable<object[]> Storages()
    {
        foreach (var storage in new[] { StringStorage.Auto, StringStorage.Shared, StringStorage.Inline })
            yield return new object[] { storage };
    }

    private static PackageOptions Options(StringStorage storage) => new PackageOptions { StringStorage = storage };

    private static List<object?[]> ValidateAndRead(byte[] workbook, string sheet = "S")
    {
        using (var doc = SpreadsheetDocument.Open(new MemoryStream(workbook), false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        }
        using var reader = ExcelReader.Open(new MemoryStream(workbook), false);
        return reader.ReadRows(sheet).Select(r => r.Values.ToArray()).ToList();
    }

    private static void AssertPlainMessage(ArgumentException e, int length)
    {
        Assert.Contains("32,767", e.Message);
        Assert.Contains(length.ToString("N0", System.Globalization.CultureInfo.InvariantCulture), e.Message);
    }

    [Theory]
    [MemberData(nameof(Storages))]
    public void WriteCell_TheLongestTextExcelOpens_IsWritten_AndOneMoreIsRefused(StringStorage storage)
    {
        var longest = new string('w', Max);
        var tooLong = new string('w', Max + 1);
        long tableBefore = -1, tableAfter = -1;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            var r = s.Row().WriteCell("first").WriteCell(longest);
            tableBefore = wb.SharedStrings.UniqueCount;
            var e = Assert.Throws<ArgumentException>(() => r.WriteCell(tooLong));
            AssertPlainMessage(e, Max + 1);
            Assert.Throws<ArgumentException>(() => r.WriteCell((object)tooLong));
            Assert.Throws<ArgumentException>(() => r.WriteInlineString(tooLong));
            tableAfter = wb.SharedStrings.UniqueCount;
            Assert.Equal(3, r.NextColumn); // The refused calls moved nothing.
            r.WriteCell("last");
            for (var i = 0; i < 300; i++) s.Row().WriteCell("x" + i).WriteCell(longest);
        }, Options(storage));

        Assert.Equal(tableBefore, tableAfter); // And added nothing to the table.
        var rows = ValidateAndRead(bytes);
        Assert.Equal(301, rows.Count);
        Assert.Equal(new object?[] { "first", longest, "last" }, rows[0]);
        Assert.Equal(longest, rows[300][1]);
    }

    /// <summary>Excel counts what the cell holds, not what the file spells: a control character is one character although it is written as <c>_x0001_</c>.</summary>
    [Theory]
    [MemberData(nameof(Storages))]
    public void TheLengthCountedIsTheTextsOwn_NotTheEscapedForm(StringStorage storage)
    {
        var text = string.Concat(Enumerable.Repeat("a\u0001_x0041_&", Max / 11)); // 11 characters each, 7 more when escaped in the file.
        text += new string('z', Max - text.Length);
        Assert.Equal(Max, text.Length);
        var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S").Row().WriteCell(text), Options(storage));
        Assert.Equal(text, ValidateAndRead(bytes)[0][0]);

        Assert.Throws<ArgumentException>(() => ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S").Row().WriteCell(text + "z"), Options(storage)));
    }

    [Fact]
    public void WriteRichText_CountsTheRunsTogether()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            var r = s.Row();
            r.WriteRichText(new RichText(new string('a', 20_000)) { Bold = true }, new RichText(new string('b', Max - 20_000)));
            var e = Assert.Throws<ArgumentException>(() => r.WriteRichText(new RichText(new string('a', 20_000)), new RichText(new string('b', Max - 20_000 + 1))));
            AssertPlainMessage(e, Max + 1);
            Assert.Equal(2, r.NextColumn);
            Assert.Throws<ArgumentException>(() => r.WriteCell((object)new[] { new RichText(new string('c', Max + 1)) }));
            r.WriteCell("after");
        });
        var row = ValidateAndRead(bytes)[0];
        Assert.Equal(new string('a', 20_000) + new string('b', Max - 20_000), row[0]);
        Assert.Equal("after", row[1]);
    }

    [Fact]
    public void WriteFormula_CachedTextOfMoreThanTheLimit_IsRefused()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            var r = s.Row();
            r.WriteFormula("REPT(\"w\",32767)", new string('w', Max));
            var e = Assert.Throws<ArgumentException>(() => r.WriteFormula("REPT(\"w\",32768)", new string('w', Max + 1)));
            AssertPlainMessage(e, Max + 1);
            Assert.Equal(2, r.NextColumn);
            r.WriteCell("after");
        });
        Assert.Equal("after", ValidateAndRead(bytes)[0][1]);
    }

    /// <summary>The link is registered before its text is written, so a refused text must not leave a link on a cell that was never written.</summary>
    [Fact]
    public void WriteHyperlink_WithTooLongText_LeavesNoLinkBehind()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            var r = s.Row();
            Assert.Throws<ArgumentException>(() => r.WriteHyperlink(new string('w', Max + 1), "https://example.com"));
            Assert.Throws<ArgumentException>(() => r.WriteInternalLink(new string('w', Max + 1), "S!A1"));
            r.WriteCell("plain");
        });
        using var zip = new System.IO.Compression.ZipArchive(new MemoryStream(bytes), System.IO.Compression.ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open());
        Assert.DoesNotContain("<hyperlink", reader.ReadToEnd());
    }

    [Fact]
    public void CellsIndexer_RefusesTooLongText_AndKeepsTheCellAsItWas()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Cells[1, 1].Value = "keep";
            var e = Assert.Throws<ArgumentException>(() => s.Cells[1, 1].Value = new string('w', Max + 1));
            AssertPlainMessage(e, Max + 1);
            Assert.Throws<ArgumentException>(() => s.Cells[2, 1].Value = new[] { new RichText(new string('w', Max + 1)) });
            Assert.Equal("keep", s.Cells[1, 1].Value);
            Assert.Single(s.Cells.Existing); // The refused set did not create a cell.
            s.Cells[2, 1].Value = new string('w', Max);
        });
        var rows = ValidateAndRead(bytes);
        Assert.Equal("keep", rows[0][0]);
        Assert.Equal(Max, ((string)rows[1][0]!).Length);
    }

    [Fact]
    public void TemplateAndEditPaths_FollowTheSameRule()
    {
        var original = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteCell("title").WriteCell(1);
        });

        // Streaming into a template sheet.
        var filled = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(original), filled, leaveOpen: true))
        {
            var s = package.Workbook.Worksheets.Open("S");
            Assert.Throws<ArgumentException>(() => s.Row(2).WriteCell(new string('w', Max + 1)));
            s.Row(2).WriteCell(new string('w', Max));
            package.Save();
        }
        Assert.Equal(Max, ((string)ValidateAndRead(filled.ToArray())[1][0]!).Length);

        // Editing cells of an existing sheet.
        var edited = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(original), edited, leaveOpen: true))
        {
            var s = package.Workbook.Worksheets.Edit("S");
            Assert.Throws<ArgumentException>(() => s.Cells["A1"].Value = new string('w', Max + 1));
            Assert.Equal("title", s.Cells["A1"].Value);
            s.Cells["A1"].Value = new string('w', Max);
            package.Save();
        }
        Assert.Equal(Max, ((string)ValidateAndRead(edited.ToArray())[0][0]!).Length);
    }

    [Fact]
    public void ObjectMapping_FollowsTheSameRule()
    {
        var rows = new[] { new Item { Name = "short" }, new Item { Name = new string('w', Max + 1) } };
        var e = Assert.Throws<ArgumentException>(() => ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S").LoadFromCollection(rows)));
        AssertPlainMessage(e, Max + 1);
    }

    private sealed class Item
    {
        public string Name { get; set; } = "";
    }
}
