using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Reading;

namespace SheetStream.Tests;

/// <summary>
/// The cached result of a formula cell. A number Excel cannot store (NaN, an infinity) is not a result at all, so
/// the cell keeps its formula and stores no result, on every build: before, the .NET build wrote <c>NaN</c> and
/// <c>Infinity</c> and the .NET Standard build wrote <c>NaN</c> and <c>INF</c>, and Excel accepts none of them.
/// </summary>
public class FormulaCachedValueTests
{
    public static IEnumerable<object[]> NotANumber() => new[]
    {
        new object[] { double.NaN },
        new object[] { double.PositiveInfinity },
        new object[] { double.NegativeInfinity },
    };

    private static byte[] Build(double cached) => ExcelPackage.CreateBytes(wb =>
        wb.Worksheets.Add("S").Row().WriteCell(1).WriteFormula("1/0", cached).WriteCell(2));

    [Theory]
    [MemberData(nameof(NotANumber))]
    public void FormulaWithNaNOrInfinityAsResult_KeepsTheFormulaAndStoresNoResult(double cached)
    {
        var bytes = Build(cached);

        using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        using var stream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
        var cells = XDocument.Load(stream).Descendants(ns + "c").ToList();

        // The cells on both sides are untouched and the formula cell is in its place with its formula.
        Assert.Equal(3, cells.Count);
        var formulaCell = cells[1];
        Assert.Equal("1/0", formulaCell.Element(ns + "f")!.Value);
        Assert.Null(formulaCell.Element(ns + "v"));
        Assert.Null(formulaCell.Attribute("t"));
        Assert.Equal("2", cells[2].Element(ns + "v")!.Value);
    }

    [Theory]
    [MemberData(nameof(NotANumber))]
    public void FormulaWithNaNOrInfinityAsResult_PassesSchemaValidationAndReadsBackWithoutAValue(double cached)
    {
        var bytes = Build(cached);

        using (var ms = new MemoryStream(bytes))
        using (var doc = SpreadsheetDocument.Open(ms, false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
            Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        }

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var row = reader.ReadRows("S").Single();
        Assert.Equal(1d, row[0]);
        Assert.Null(row[1]);
        Assert.Equal(2d, row[2]);
    }

    [Fact]
    public void FormulaWithARealResult_StillStoresIt()
    {
        using var reader = ExcelReader.Open(new MemoryStream(Build(0.5)));
        Assert.Equal(0.5, reader.ReadRows("S").Single()[1]);
    }
}
