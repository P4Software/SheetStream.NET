using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream.Tests;

/// <summary>Unit checks on the forward-only contract, validation rules and helper logic.</summary>
public class ApiContractTests
{
    private static ExcelPackage NewPackage() => new ExcelPackage(new MemoryStream(), leaveOpen: false);

    [Theory]
    [InlineData(1, "A")]
    [InlineData(26, "Z")]
    [InlineData(27, "AA")]
    [InlineData(52, "AZ")]
    [InlineData(53, "BA")]
    [InlineData(702, "ZZ")]
    [InlineData(703, "AAA")]
    [InlineData(16384, "XFD")]
    public void ColumnLetters_MatchExcel(int index, string expected)
        => Assert.Equal(expected, CellReference.ColumnLetters(index));

    [Fact]
    public void Sanitize_EscapesControlCharsLikeExcel()
    {
        Assert.Equal("a_x0000_b", XmlTextUtil.Sanitize("a\0b"));
        Assert.Equal("a_x000C_b", XmlTextUtil.Sanitize("a\fb"));
        Assert.Equal("tab\tnl\ncr\r", XmlTextUtil.Sanitize("tab\tnl\ncr\r"));
        Assert.Equal("_x005F_x0041_", XmlTextUtil.Sanitize("_x0041_"));
        Assert.Equal("plain_x41_", XmlTextUtil.Sanitize("plain_x41_"));
        Assert.Equal("�", XmlTextUtil.Sanitize("\uD800"));
        var same = "no changes needed";
        Assert.Same(same, XmlTextUtil.Sanitize(same));
    }

    [Theory]
    [InlineData("#4F81BD", "FF4F81BD")]
    [InlineData("4f81bd", "FF4F81BD")]
    [InlineData("804F81BD", "804F81BD")]
    [InlineData("#804F81BD", "804F81BD")]
    public void NormalizeColor_ProducesArgb(string input, string expected)
        => Assert.Equal(expected, XmlTextUtil.NormalizeColor(input, "c"));

    [Theory]
    [InlineData("#4F81B")]
    [InlineData("GGGGGG")]
    [InlineData("")]
    public void NormalizeColor_RejectsBadInput(string input)
        => Assert.Throws<ArgumentException>(() => XmlTextUtil.NormalizeColor(input, "c"));

    [Fact]
    public void Styles_AreDeduplicated()
    {
        using var package = NewPackage();
        var a = package.Workbook.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "0.00"; });
        var b = package.Workbook.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "0.00"; });
        var c = package.Workbook.Styles.Create(o => { o.Font.Bold = true; o.NumberFormat = "0.000"; });
        Assert.Equal(a.Id, b.Id);
        Assert.NotEqual(a.Id, c.Id);
        Assert.Equal(0, package.Workbook.Styles.Default.Id);
    }

    [Fact]
    public void ColumnSettings_AfterFirstRow_Throw()
    {
        using var package = NewPackage();
        var sheet = package.Workbook.Worksheets.Add("S");
        sheet.Row().WriteCell(1);
        var ex = Assert.Throws<InvalidOperationException>(() => sheet.Column(1).SetWidth(10));
        Assert.Contains("before the first row", ex.Message);
        Assert.Throws<InvalidOperationException>(() => sheet.FreezePanes(1));
    }

    [Fact]
    public void Rows_MustBeAscending()
    {
        using var package = NewPackage();
        var sheet = package.Workbook.Worksheets.Add("S");
        sheet.Row(5).WriteCell("five");
        Assert.Same(sheet.Row(5), sheet.Row(5)); // Re-entering the open row is allowed.
        sheet.Row(5).WriteCell("still five");
        Assert.Throws<InvalidOperationException>(() => sheet.Row(4));
        sheet.Row(10).WriteCell("ten");
        Assert.Equal(10, sheet.CurrentRow);
    }

    [Fact]
    public void AddingSheet_SealsPrevious()
    {
        using var package = NewPackage();
        var first = package.Workbook.Worksheets.Add("One");
        first.Row().WriteCell(1);
        var second = package.Workbook.Worksheets.Add("Two");
        Assert.True(first.IsSealed);
        Assert.False(second.IsSealed);
        Assert.Throws<InvalidOperationException>(() => first.Row().WriteCell(2));
        Assert.Same(second, package.Workbook.Worksheets.Current);
    }

    [Theory]
    [InlineData("")]
    [InlineData("ThisSheetNameIsLongerThanThirtyOneChars")]
    [InlineData("bad:name")]
    [InlineData("bad/name")]
    [InlineData("bad[name]")]
    [InlineData("'quoted")]
    public void SheetNames_AreValidated(string name)
    {
        using var package = NewPackage();
        Assert.Throws<ArgumentException>(() => package.Workbook.Worksheets.Add(name));
    }

    [Fact]
    public void SheetNames_MustBeUniqueIgnoringCase()
    {
        using var package = NewPackage();
        package.Workbook.Worksheets.Add("Data");
        Assert.Throws<ArgumentException>(() => package.Workbook.Worksheets.Add("DATA"));
    }

    [Fact]
    public void WritingAfterSave_Throws()
    {
        using var package = NewPackage();
        var sheet = package.Workbook.Worksheets.Add("S");
        package.Save();
        Assert.True(package.IsSaved);
        Assert.Throws<InvalidOperationException>(() => sheet.Row().WriteCell(1));
        Assert.Throws<InvalidOperationException>(() => package.Workbook.Worksheets.Add("T"));
    }

    [Fact]
    public void WriteCell_WithoutRow_Throws()
    {
        using var package = NewPackage();
        var sheet = package.Workbook.Worksheets.Add("S");
        var row = sheet.Row();
        sheet.Workbook.Worksheets.Add("T"); // seals S, closing its row
        Assert.Throws<InvalidOperationException>(() => row.WriteCell(1));
    }

    [Fact]
    public void SharedStrings_CountReferencesAndUniques()
    {
        using var package = NewPackage();
        var sheet = package.Workbook.Worksheets.Add("S");
        sheet.Row().WriteCell("a").WriteCell("b").WriteCell("a");
        Assert.Equal(2, package.Workbook.SharedStrings.UniqueCount);
        Assert.Equal(3, package.Workbook.SharedStrings.ReferenceCount);
    }

    [Fact]
    public void BorderStyle_SetsAllEdges()
    {
        var o = new StyleOptions();
        o.Border.Style = BorderStyle.Medium;
        Assert.Equal(BorderStyle.Medium, o.Border.Left);
        Assert.Equal(BorderStyle.Medium, o.Border.Right);
        Assert.Equal(BorderStyle.Medium, o.Border.Top);
        Assert.Equal(BorderStyle.Medium, o.Border.Bottom);
        Assert.True(o.Border.HasAnyLine);
    }

    [Fact]
    public void Colors_Rgb_FormatsHex()
        => Assert.Equal("#4F81BD", Colors.Rgb(79, 129, 189));
}
