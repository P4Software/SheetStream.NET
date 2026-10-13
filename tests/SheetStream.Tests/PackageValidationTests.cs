// The Open XML SDK exposes nearly every member as nullable; in test assertions a null is a failure anyway.
#pragma warning disable CS8602
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Styles;

namespace SheetStream.Tests;

/// <summary>
/// End-to-end checks: build a workbook, validate it against the ECMA-376 schema with the Open XML SDK
/// validator, then read it back and assert the values landed where they should.
/// </summary>
public class PackageValidationTests
{
    /// <summary>Builds a workbook exercising every feature and returns its bytes.</summary>
    private static byte[] BuildFullFeaturedWorkbook()
    {
        return ExcelPackage.CreateBytes(wb =>
        {
            var sheet = wb.Worksheets.Add("Data");
            sheet.Column(1).SetWidth(10);
            sheet.Column(2).SetWidth(30).Hide();
            sheet.Column(3).SetWidth(18);
            sheet.FreezePanes(1, 1);
            sheet.AutoFilter(1, 1, 3);

            var header = wb.Styles.Create(o =>
            {
                o.Font.Bold = true;
                o.Font.Italic = true;
                o.Font.Color = "#FFFFFF";
                o.Fill.BackgroundColor = "#4F81BD";
                o.Border.Style = BorderStyle.Thin;
                o.Border.Color = "#000000";
                o.HorizontalAlignment = HorizontalAlignment.Center;
                o.VerticalAlignment = VerticalAlignment.Center;
                o.WrapText = true;
            });
            var money = wb.Styles.Create(o => o.NumberFormat = "$#,##0.00");
            var pct = wb.Styles.Create(o => o.NumberFormat = "0.00%");

            sheet.Row(1).WriteCell("Id", header).WriteCell("Name", header).WriteCell("Amount", header);
            sheet.Row(2).WriteCell(1).WriteCell("Alpha").WriteCell(10.5, money);
            sheet.Row(3).WriteCell(2).WriteCell("Beta").WriteCell(20.25m, money);
            sheet.Row(4).WriteCell(3).WriteCell("Alpha").WriteCell(0.5, pct);
            sheet.Row(6, null, 30).WriteCell("Gap row with height").Skip(1).WriteCell(true);
            sheet.Row(7).WriteCell(new DateTime(2026, 10, 13)).WriteCell(new DateTime(2026, 10, 13, 15, 30, 0)).WriteCell(TimeSpan.FromHours(26));
            sheet.Row(8).WriteCell(double.NaN).WriteInlineString("  padded  ").WriteCell("Line\nbreak \u0001 control");
            sheet.Row(9).WriteCell("Total", header).WriteEmpty(header).WriteFormula("=SUM(C2:C4)", money);
            sheet.MergeCells(10, 1, 10, 3);
            sheet.Row(10).WriteCell("Merged across three columns");
            var stripe = wb.Styles.Create(o => o.Fill.BackgroundColor = "#F2F2F2");
            sheet.Row(11, stripe).WriteCell(new DateTime(2026, 1, 2)).WriteCell(new DateTime(2026, 1, 2, 8, 0, 0)).WriteCell(TimeSpan.FromMinutes(90)).WriteCell("text");

            var second = wb.Worksheets.Add("Empty");
            _ = second;

            var third = wb.Worksheets.Add("Mixed");
            third.Row().WriteCells("x", 1, 2L, 3.5f, 4.5m, (byte)5, 'c', Guid.Empty, DayOfWeek.Monday, null, DBNull.Value, new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero));
        });
    }

    [Fact]
    public void FullFeaturedWorkbook_PassesOpenXmlSchemaValidation()
    {
        var bytes = BuildFullFeaturedWorkbook();

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);
        var validator = new OpenXmlValidator(FileFormatVersions.Office2016);
        var errors = validator.Validate(doc).ToList();

        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
    }

    [Fact]
    public void FullFeaturedWorkbook_RoundTripsValues()
    {
        var bytes = BuildFullFeaturedWorkbook();

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);
        var wbPart = doc.WorkbookPart!;
        var sheets = wbPart.Workbook.Sheets!.Elements<Sheet>().ToList();
        Assert.Equal(new[] { "Data", "Empty", "Mixed" }, sheets.Select(s => s.Name!.Value));

        var sst = wbPart.SharedStringTablePart!.SharedStringTable;
        var strings = sst.Elements<SharedStringItem>().Select(si => si.InnerText).ToList();
        // "Alpha" appears twice in the data but once in the table.
        Assert.Single(strings, s => s == "Alpha");
        Assert.Contains("Line\nbreak _x0001_ control", strings);

        var dataPart = (WorksheetPart)wbPart.GetPartById(sheets[0].Id!.Value!);
        var cells = dataPart.Worksheet.Descendants<Cell>().ToDictionary(c => c.CellReference!.Value!, c => c);

        Assert.Equal(CellValues.SharedString, cells["A1"].DataType!.Value);
        Assert.Equal("Id", strings[int.Parse(cells["A1"].CellValue!.Text)]);
        Assert.NotNull(cells["A1"].StyleIndex);
        Assert.Equal("1", cells["A2"].CellValue!.Text);
        Assert.Equal("10.5", cells["C2"].CellValue!.Text);
        Assert.Equal("20.25", cells["C3"].CellValue!.Text);
        Assert.Equal(CellValues.Boolean, cells["C6"].DataType!.Value);
        Assert.Equal("1", cells["C6"].CellValue!.Text);
        Assert.Equal(new DateTime(2026, 10, 13).ToOADate().ToString("R"), cells["A7"].CellValue!.Text);
        Assert.Equal(CellValues.Error, cells["A8"].DataType!.Value);
        Assert.Equal("#NUM!", cells["A8"].CellValue!.Text);
        Assert.Equal(CellValues.InlineString, cells["B8"].DataType!.Value);
        Assert.Equal("  padded  ", cells["B8"].InlineString!.Text!.Text);
        Assert.Equal("SUM(C2:C4)", cells["C9"].CellFormula!.Text);
        Assert.False(cells.ContainsKey("B6")); // Skipped cell is absent, not empty.

        // Row 6 carries the custom height; row 5 was never written.
        var rows = dataPart.Worksheet.Descendants<Row>().ToList();
        Assert.DoesNotContain(rows, r => r.RowIndex!.Value == 5);
        Assert.Equal(30, rows.Single(r => r.RowIndex!.Value == 6).Height!.Value);

        Assert.Equal("A1:C1", dataPart.Worksheet.GetFirstChild<AutoFilter>()!.Reference!.Value);
        Assert.Equal("A10:C10", dataPart.Worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Single().Reference!.Value);
        var cols = dataPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToList();
        Assert.Equal(3, cols.Count);
        Assert.True(cols[1].Hidden!.Value);
        Assert.Equal(30, cols[1].Width!.Value);
        var pane = dataPart.Worksheet.GetFirstChild<SheetViews>()!.Descendants<Pane>().Single();
        Assert.Equal("B2", pane.TopLeftCell!.Value);

        Assert.NotNull(wbPart.Workbook.CalculationProperties);
        Assert.True(wbPart.Workbook.CalculationProperties!.FullCalculationOnLoad!.Value);
        Assert.Contains(wbPart.Workbook.DefinedNames!.Elements<DefinedName>(), d => d.Name!.Value == "_xlnm._FilterDatabase" && d.Text == "'Data'!$A$1:$C$1");

        // Styles: header has bold/italic font, solid fill and thin border; money has a custom numFmt.
        var styles = wbPart.WorkbookStylesPart!.Stylesheet;
        var headerXf = styles.CellFormats!.Elements<CellFormat>().ElementAt((int)cells["A1"].StyleIndex!.Value);
        var font = styles.Fonts!.Elements<Font>().ElementAt((int)headerXf.FontId!.Value);
        Assert.NotNull(font.Bold);
        Assert.NotNull(font.Italic);
        Assert.Equal("FFFFFFFF", font.Color!.Rgb!.Value);
        var fill = styles.Fills!.Elements<Fill>().ElementAt((int)headerXf.FillId!.Value);
        Assert.Equal(PatternValues.Solid, fill.PatternFill!.PatternType!.Value);
        Assert.Equal("FF4F81BD", fill.PatternFill.ForegroundColor!.Rgb!.Value);
        var border = styles.Borders!.Elements<Border>().ElementAt((int)headerXf.BorderId!.Value);
        Assert.Equal(BorderStyleValues.Thin, border.LeftBorder!.Style!.Value);
        var moneyXf = styles.CellFormats.Elements<CellFormat>().ElementAt((int)cells["C2"].StyleIndex!.Value);
        Assert.Equal("$#,##0.00", styles.NumberingFormats!.Elements<NumberingFormat>().Single(n => n.NumberFormatId!.Value == moneyXf.NumberFormatId!.Value).FormatCode!.Value);
        // Dates on a row with a default (fill-only) style keep the fill AND get a date format.
        var cellFormats = styles.CellFormats.Elements<CellFormat>().ToList();
        var stripedDate = cellFormats[(int)cells["A11"].StyleIndex!.Value];
        var stripedDateTime = cellFormats[(int)cells["B11"].StyleIndex!.Value];
        var stripedDuration = cellFormats[(int)cells["C11"].StyleIndex!.Value];
        var stripedText = cellFormats[(int)cells["D11"].StyleIndex!.Value];
        Assert.Equal(14u, stripedDate.NumberFormatId!.Value);
        Assert.Equal(22u, stripedDateTime.NumberFormatId!.Value);
        Assert.Equal(46u, stripedDuration.NumberFormatId!.Value);
        Assert.Equal(0u, stripedText.NumberFormatId!.Value);
        Assert.Equal(stripedText.FillId!.Value, stripedDate.FillId!.Value);
        Assert.Equal("FFF2F2F2", styles.Fills.Elements<Fill>().ElementAt((int)stripedDate.FillId!.Value).PatternFill!.ForegroundColor!.Rgb!.Value);

        // Fills 0 and 1 are the mandatory none / gray125 placeholders.
        var fills = styles.Fills.Elements<Fill>().ToList();
        Assert.Equal(PatternValues.None, fills[0].PatternFill!.PatternType!.Value);
        Assert.Equal(PatternValues.Gray125, fills[1].PatternFill!.PatternType!.Value);
    }

    [Fact]
    public void InlineStringMode_WritesNoSharedStrings()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            wb.Worksheets.Add("S").Row().WriteCell("a").WriteCell("b");
        }, new PackageOptions { UseSharedStrings = false });

        using var ms = new MemoryStream(bytes);
        using var doc = SpreadsheetDocument.Open(ms, false);
        Assert.Empty(doc.WorkbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>());
        var cell = doc.WorkbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>().First();
        Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc));
    }

    [Fact]
    public void DisposeWithoutSave_StillProducesCompleteFile()
    {
        using var ms = new MemoryStream();
        using (var package = new ExcelPackage(ms, leaveOpen: true))
        {
            package.Workbook.Worksheets.Add("Only").Row().WriteCell(42);
        }
        ms.Position = 0;
        using var doc = SpreadsheetDocument.Open(ms, false);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc));
        Assert.Equal("42", doc.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>().Single().CellValue!.Text);
    }

    [Fact]
    public void LargeSheet_StreamsWithoutRetainingRows()
    {
        const int rows = 200_000;
        var before = GC.GetTotalMemory(forceFullCollection: true);
        using var sink = new CountingStream();
        using (var package = new ExcelPackage(sink, leaveOpen: true, new PackageOptions { CompressionLevel = System.IO.Compression.CompressionLevel.Fastest }))
        {
            var sheet = package.Workbook.Worksheets.Add("Big");
            for (var i = 1; i <= rows; i++)
                sheet.Row().WriteCell(i).WriteCell(i * 0.5).WriteCell(i % 10 == 0 ? "ten" : "other");
        }
        var after = GC.GetTotalMemory(forceFullCollection: true);

        Assert.True(sink.Length > 100_000, "Output was written.");
        // 200k rows x 3 cells would be tens of MB in a DOM. Streaming leaves nothing behind beyond buffers.
        Assert.True(after - before < 8 * 1024 * 1024, $"Retained {(after - before) / 1024} KB after writing {rows} rows.");
    }

    /// <summary>Write-only stream that counts bytes, so the large-sheet test does not need disk or memory for the output.</summary>
    private sealed class CountingStream : Stream
    {
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        private long _length;
        public override long Length => _length;
        public override long Position { get => _length; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => _length += count;
    }
}
