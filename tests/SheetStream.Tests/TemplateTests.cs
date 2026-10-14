#pragma warning disable CS8602, CS8603, CS8604, CS8714, CS8621
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Reading;
using SheetStream.Styles;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Tests;

/// <summary>Filling a template: copy everything, stream rows into the data area, extend the template's ranges.</summary>
public class TemplateTests
{
    /// <summary>
    /// A template like a designer would make: title, header row 4 as a table with one formatted data row 5,
    /// a validation and a conditional format on that row, a SUM over the table, and an untouched second sheet.
    /// </summary>
    private static byte[] BuildTemplate()
    {
        return ExcelPackage.CreateBytes(wb =>
        {
            var title = wb.Styles.Create(o => { o.Font.Bold = true; o.Font.Size = 18; });
            var header = wb.Styles.Create(o => { o.Font.Bold = true; o.Fill.BackgroundColor = "#0A5FAD"; o.Font.Color = "#FFFFFF"; });
            var money = wb.Styles.Create(o => o.NumberFormat = "$#,##0.00");
            var red = wb.Styles.CreateDifferential(o => o.Font.Color = "#C00000");
            var s = wb.Worksheets.Add("Lines");
            s.Column(2).SetWidth(30);
            s.FreezePanes(4);
            s.AddTable("Lines", "A4:D5");
            s.AddDataValidation("C5").WholeNumber(ValidationOperator.GreaterThan, 0);
            s.AddConditionalFormatting("D5").CellValue(ConditionalOperator.LessThan, "0", red);
            s.Row(1, title).WriteCell("Invoice");
            s.Row(2).WriteCell("Total").Skip(2).WriteFormula("SUM(Lines[Amount])", money);
            s.Row(4).WriteCell("Sku", header).WriteCell("Description", header).WriteCell("Qty", header).WriteCell("Amount", header);
            s.Row(5).WriteEmpty().WriteEmpty().WriteEmpty().WriteEmpty(money);
            s.Row(8).WriteCell("Thank you for your business"); // A footer the data will overrun.
            var notes = wb.Worksheets.Add("Notes");
            notes.Row().WriteRichText("Keep ", new RichText("this") { Bold = true });
            notes.AddPicture(FeatureTests.TinyPng, 2, 2);
        });
    }

    private static SpreadsheetDocument OpenValidated(byte[] bytes)
    {
        var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        return doc;
    }

    private static byte[] Fill(byte[] template, Action<ExcelPackage> fill)
    {
        using var output = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true))
        {
            fill(package);
            package.Save();
        }
        return output.ToArray();
    }

    [Fact]
    public void FillingTemplate_KeepsLayoutAndExtendsRanges()
    {
        var template = BuildTemplate();
        int replaced = -1;
        var filled = Fill(template, p =>
        {
            var s = p.Workbook.Worksheets.Open("lines");     // Case-insensitive.
            var money = s.TemplateStyle("D5");
            var stripe = p.Workbook.Styles.Create(o => o.Fill.BackgroundColor = "#F2F2F2"); // A new style on top of the template's.
            for (var i = 0; i < 10; i++)
                s.Row(5 + i).WriteCell("SKU-" + i).WriteCell(i % 2 == 0 ? "Widget" : "Gadget", i % 2 == 0 ? stripe : null).WriteCell(i + 1).WriteCell((i + 1) * 9.5, money);
            replaced = s.ReplacedTemplateRows;
        });
        Assert.Equal(1, replaced); // Row 8's footer was overrun by the data.

        using var doc = OpenValidated(filled);
        var book = doc.WorkbookPart;
        Assert.Equal(new[] { "Lines", "Notes" }, book.Workbook.Sheets.Elements<X.Sheet>().Select(x => x.Name.Value));
        Assert.True(book.Workbook.CalculationProperties.FullCalculationOnLoad.Value);

        var linesPart = (WorksheetPart)book.GetPartById(book.Workbook.Sheets.Elements<X.Sheet>().First().Id);
        var ws = linesPart.Worksheet;
        Assert.Null(ws.SheetDimension);                                          // Stale hint dropped.
        Assert.NotNull(ws.SheetViews);                                            // Frozen pane kept.
        Assert.Equal(30, ws.GetFirstChild<X.Columns>().Elements<X.Column>().Single().Width.Value);
        var table = linesPart.TableDefinitionParts.Single().Table;
        Assert.Equal("A4:D14", table.Reference.Value);                            // Extended over the 10 rows.
        Assert.Equal("A4:D14", table.AutoFilter.Reference.Value);
        Assert.Equal("C5:C14", ws.GetFirstChild<X.DataValidations>().Elements<X.DataValidation>().Single().SequenceOfReferences.InnerText);
        Assert.Equal("D5:D14", ws.Elements<X.ConditionalFormatting>().Single().SequenceOfReferences.InnerText);

        // Title and formula rows above the data survive; template styles are still applied.
        var cells = ws.Descendants<X.Cell>().ToDictionary(c => c.CellReference.Value);
        Assert.Equal("SUM(Lines[Amount])", cells["D2"].CellFormula.Text);
        Assert.NotNull(cells["A1"].StyleIndex);
        Assert.Equal(cells["D14"].StyleIndex.Value, cells["D5"].StyleIndex.Value);
        Assert.False(cells.ContainsKey("A8") && cells["A8"].DataType?.Value == X.CellValues.SharedString
            && book.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cells["A8"].CellValue.Text)).InnerText.StartsWith("Thank"));

        // The untouched sheet and its picture are copied as they were.
        var notesPart = (WorksheetPart)book.GetPartById(book.Workbook.Sheets.Elements<X.Sheet>().Last().Id);
        Assert.Single(notesPart.DrawingsPart.ImageParts);

        // The new fill style was appended after the template's styles, and template fills are intact.
        var fills = book.WorkbookStylesPart.Stylesheet.Fills.Elements<X.Fill>().ToList();
        Assert.Contains(fills, f => f.PatternFill?.ForegroundColor?.Rgb?.Value == "FF0A5FAD");
        Assert.Contains(fills, f => f.PatternFill?.ForegroundColor?.Rgb?.Value == "FFF2F2F2");
        Assert.Equal((uint)fills.Count, book.WorkbookStylesPart.Stylesheet.Fills.Count.Value);

        using var reader = ExcelReader.Open(new MemoryStream(filled));
        var rows = reader.ReadRows("Lines").ToList();
        Assert.Equal("Invoice", rows[0][0]);
        Assert.Equal("Sku", rows.Single(r => r.Index == 4)[0]);
        Assert.Equal("SKU-9", rows.Single(r => r.Index == 14)[0]);
        Assert.Equal(95.0, rows.Single(r => r.Index == 14)[3]);
        Assert.Equal("Keep this", reader.ReadRows("Notes").First()[0]);   // Rich text in the template's string table survived.
    }

    [Fact]
    public void Template_GapsAndFootersBelowDataAreKept()
    {
        var filled = Fill(BuildTemplate(), p =>
        {
            var s = p.Workbook.Worksheets.Open("Lines");
            s.Row(5).WriteCell("only one");
        });
        using var reader = ExcelReader.Open(new MemoryStream(filled));
        var rows = reader.ReadRows("Lines").ToDictionary(r => r.Index);
        Assert.Equal("only one", rows[5][0]);
        Assert.Equal("Thank you for your business", rows[8][0]);
        using var doc = OpenValidated(filled);
    }

    [Fact]
    public void Template_ReusesTemplateStringsAndAddsDateStyles()
    {
        var filled = Fill(BuildTemplate(), p =>
        {
            p.Workbook.Worksheets.Open("Lines").Row(5).WriteCell("Description").WriteCell(new DateTime(2026, 10, 13));
        });
        using var doc = OpenValidated(filled);
        var sst = doc.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<X.SharedStringItem>().Select(i => i.InnerText).ToList();
        Assert.Single(sst, t => t == "Description");    // Reused the template's string instead of adding a duplicate.
        using var reader = ExcelReader.Open(new MemoryStream(filled));
        Assert.Equal(new DateTime(2026, 10, 13), reader.ReadRows("Lines").Single(r => r.Index == 5)[1]);
    }

    [Fact]
    public void Template_RejectsLayoutChangesAndUnknownSheets()
    {
        using var output = new MemoryStream();
        using var p = ExcelPackage.OpenTemplate(new MemoryStream(BuildTemplate()), output, leaveOpen: true);
        Assert.Throws<NotSupportedException>(() => p.Workbook.Worksheets.Add("New"));
        Assert.Throws<ArgumentException>(() => p.Workbook.Worksheets.Open("Missing"));
        var s = p.Workbook.Worksheets.Open("Lines");
        Assert.Throws<NotSupportedException>(() => s.Column(1).SetWidth(5));
        Assert.Throws<NotSupportedException>(() => s.AutoFitColumns());
        Assert.Throws<InvalidOperationException>(() => p.Workbook.Worksheets.Open("Lines"));
    }

    [Fact]
    public void OpenTemplate_RefusesToOverwriteTheTemplate()
    {
        var path = Path.Combine(Path.GetTempPath(), "ss-template-" + Guid.NewGuid() + ".xlsx");
        File.WriteAllBytes(path, BuildTemplate());
        try { Assert.Throws<ArgumentException>(() => ExcelPackage.OpenTemplate(path, path)); }
        finally { File.Delete(path); }
    }
}
