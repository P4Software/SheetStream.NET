using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// An XML parser turns a raw carriage return into a line feed (XML 1.0, section 2.11), so text written with one came back
/// as <c>"a\nb"</c> for <c>"a\r\nb"</c>. The character reference <c>&amp;#xD;</c> survives parsing and is read back as a CR
/// by Excel (checked through COM), ExcelReader and openpyxl; Excel's own <c>_x000D_</c> is decoded by Excel and
/// ExcelReader but shown literally by openpyxl, so it is not used. Text without a CR is written as it was.
/// </summary>
public class CarriageReturnTests
{
    private static readonly string[] Texts =
    {
        "a\r\nb", "a\rb", "\r", "\r\n", "\r\r\n\n\r", " lead\r\ntrail\r\n ", "x & <y>\r\n\"q\"\r_x0041_\t\u0001", "é€😀\r\nz",
    };

    public static IEnumerable<object[]> Storages()
    {
        foreach (var storage in new[] { StringStorage.Auto, StringStorage.Shared, StringStorage.Inline })
            yield return new object[] { storage };
    }

    private static string Part(byte[] workbook, string name)
    {
        using var zip = new ZipArchive(new MemoryStream(workbook), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry(name)!.Open(), new UTF8Encoding(false));
        return reader.ReadToEnd();
    }

    private static List<object?[]> Read(byte[] workbook)
    {
        using var reader = ExcelReader.Open(new MemoryStream(workbook), false);
        return reader.ReadRows("S").Select(r => r.Values.ToArray()).ToList();
    }

    private static void AssertValid(byte[] workbook)
    {
        using var doc = SpreadsheetDocument.Open(new MemoryStream(workbook), false);
        Assert.Empty(new OpenXmlValidator().Validate(doc));
    }

    [Theory]
    [MemberData(nameof(Storages))]
    public void CarriageReturns_SurviveTheRoundTrip_InEveryStorageMode(StringStorage storage)
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            foreach (var text in Texts) s.Row().WriteCell(text).WriteInlineString(text);
        }, new PackageOptions { StringStorage = storage });

        var rows = Read(bytes);
        Assert.Equal(Texts.Length, rows.Count);
        for (var i = 0; i < Texts.Length; i++)
        {
            Assert.Equal(Texts[i], rows[i][0]);
            Assert.Equal(Texts[i], rows[i][1]);
        }
        // Not one raw CR in the parts that hold text: they are all character references.
        Assert.DoesNotContain('\r', Part(bytes, "xl/worksheets/sheet1.xml"));
        if (storage != StringStorage.Inline) Assert.DoesNotContain('\r', Part(bytes, "xl/sharedStrings.xml"));
        AssertValid(bytes);
    }

    /// <summary>The shared string table composes short texts by hand and sends long ones through the XML writer: both must agree.</summary>
    [Fact]
    public void CarriageReturns_InALongSharedString_AreWrittenTheSameWay()
    {
        var longText = string.Concat(Enumerable.Range(0, 3000).Select(i => "line " + i + "\r\n"));
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteCell(longText).WriteCell("short\r\ntext");
        }, new PackageOptions { StringStorage = StringStorage.Shared });

        var row = Read(bytes)[0];
        Assert.Equal(longText, row[0]);
        Assert.Equal("short\r\ntext", row[1]);
        Assert.DoesNotContain('\r', Part(bytes, "xl/sharedStrings.xml"));
    }

    [Fact]
    public void CarriageReturns_InRichTextAndFormulaResults_AreKept()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteRichText(new RichText("one\r\ntwo") { Bold = true }, new RichText("\rthree"))
                   .WriteFormula("\"a\"&CHAR(13)&CHAR(10)&\"b\"", "a\r\nb");
        });
        var row = Read(bytes)[0];
        Assert.Equal("one\r\ntwo\rthree", row[0]);
        Assert.Equal("a\r\nb", row[1]);
        Assert.DoesNotContain('\r', Part(bytes, "xl/worksheets/sheet1.xml"));
        AssertValid(bytes);
    }

    /// <summary>A workbook written here is a fine template or edit source: the text it holds must come out the other side unchanged.</summary>
    [Fact]
    public void CarriageReturns_SurviveTemplateAndEditCopies()
    {
        var original = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteCell("keep\r\nme").WriteCell(1);
        }, new PackageOptions { StringStorage = StringStorage.Shared });

        var filled = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(original), filled, leaveOpen: true))
        {
            package.Workbook.Worksheets.Open("S").Row(2).WriteCell("new\r\ntext");
            package.Save();
        }
        var rows = Read(filled.ToArray());
        Assert.Equal("keep\r\nme", rows[0][0]);
        Assert.Equal("new\r\ntext", rows[1][0]);

        var edited = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(original), edited, leaveOpen: true))
        {
            package.Workbook.Worksheets.Edit("S").Cells["B1"].Value = "edited\r\nvalue";
            package.Save();
        }
        rows = Read(edited.ToArray());
        Assert.Equal("keep\r\nme", rows[0][0]);
        Assert.Equal("edited\r\nvalue", rows[0][1]);
    }

    /// <summary>The bytes of text without a carriage return are what they were.</summary>
    [Fact]
    public void TextWithoutACarriageReturn_IsWrittenAsBefore()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteInlineString("a\nb\tc & <d>").WriteCell("e\nf");
        }, new PackageOptions { StringStorage = StringStorage.Shared });
        Assert.Contains("<t>a\nb\tc &amp; &lt;d&gt;</t>", Part(bytes, "xl/worksheets/sheet1.xml"));
        Assert.Contains("<t>e\nf</t>", Part(bytes, "xl/sharedStrings.xml"));
    }
}
