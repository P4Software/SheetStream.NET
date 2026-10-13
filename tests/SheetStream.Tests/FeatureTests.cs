// The Open XML SDK exposes nearly every member as nullable; in test assertions a null is a failure anyway.
#pragma warning disable CS8602, CS8603, CS8604
using System.ComponentModel;
using System.Data;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Reading;
using SheetStream.Styles;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Tests;

/// <summary>
/// Covers the EPPlus-parity features: every test builds a workbook, validates it against the ECMA-376 schema
/// and checks the written structure or reads it back.
/// </summary>
public class FeatureTests
{
    /// <summary>A 2x1 red PNG, small enough to inline.</summary>
    internal static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAIAAAABCAIAAAB7QOjdAAAAEElEQVR4nGP4z8DAwMDAAAAD9QH/mKGklwAAAABJRU5ErkJggg==");

    private static SpreadsheetDocument OpenValidated(byte[] bytes)
    {
        var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        return doc;
    }

    private static X.Worksheet SheetXml(SpreadsheetDocument doc, string name)
    {
        var sheet = doc.WorkbookPart.Workbook.Sheets.Elements<X.Sheet>().Single(s => s.Name == name);
        return ((WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id)).Worksheet;
    }

    [Fact]
    public void ConditionalFormattingAndValidation_AreWrittenAndValid()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var bad = wb.Styles.CreateDifferential(o => { o.Font.Color = "#9C0006"; o.Fill.BackgroundColor = "#FFC7CE"; o.Font.Bold = true; });
            var good = wb.Styles.CreateDifferential(o => { o.Fill.BackgroundColor = "#C6EFCE"; o.NumberFormat = "0.0%"; });
            var s = wb.Worksheets.Add("Rules");
            s.AddConditionalFormatting("B2:B20")
                .CellValue(ConditionalOperator.LessThan, "0", bad)
                .CellValue(ConditionalOperator.Between, "10", good, "20")
                .Expression("$B2>$C2", good)
                .ContainsText("err\"or", bad)
                .DuplicateValues(bad)
                .Top(3, good)
                .Bottom(10, bad, percent: true)
                .AboveAverage(good)
                .ColorScale("#F8696B", "#FFEB84", "#63BE7B")
                .DataBar("#638EC6")
                .IconSet(IconSetType.FiveRating, reverse: true);
            s.AddConditionalFormatting(2, 3, 20, 3).ColorScale("#FFFFFF", "#5A8AC6");
            s.AddDataValidation("D2:D20").List("Open", "Closed", "On \"hold\"").WithError("Bad", "Pick one").WithPrompt("Status", "Choose");
            s.AddDataValidation("E2:E20").WholeNumber(ValidationOperator.Between, 1, 10);
            s.AddDataValidation("F2:F20").Decimal(ValidationOperator.GreaterThan, 0.5);
            s.AddDataValidation("G2:G20").Date(ValidationOperator.LessThanOrEqual, new DateTime(2030, 1, 1));
            s.AddDataValidation("H2:H20").TextLength(ValidationOperator.LessThan, 50);
            s.AddDataValidation("I2:I20").Custom("=ISNUMBER(I2)");
            s.AddDataValidation("J2:J20").ListFromRange("'Rules'!K1:K5");
            for (var r = 1; r <= 20; r++) s.Row().WriteCell(r).WriteCell(r - 5).WriteCell(r * 2);
        });

        using var doc = OpenValidated(bytes);
        var ws = SheetXml(doc, "Rules");
        var rules = ws.Elements<X.ConditionalFormatting>().SelectMany(c => c.Elements<X.ConditionalFormattingRule>()).ToList();
        Assert.Equal(12, rules.Count);
        Assert.Equal(Enumerable.Range(1, 12), rules.Select(r => r.Priority.Value)); // Unique, ascending priorities.
        Assert.Equal("NOT(ISERROR(SEARCH(\"err\"\"or\",B2)))", rules.Single(r => r.Type == X.ConditionalFormatValues.ContainsText).Elements<X.Formula>().Single().Text);
        var dxfs = doc.WorkbookPart.WorkbookStylesPart.Stylesheet.DifferentialFormats.Elements<X.DifferentialFormat>().ToList();
        Assert.Equal(2, dxfs.Count);
        Assert.Equal("FFFFC7CE", dxfs[0].Fill.PatternFill.BackgroundColor.Rgb.Value); // dxf fills use bgColor.

        var validations = ws.GetFirstChild<X.DataValidations>().Elements<X.DataValidation>().ToList();
        Assert.Equal(7, validations.Count);
        Assert.Equal("\"Open,Closed,On \"\"hold\"\"\"", validations[0].Formula1.Text);
        Assert.Equal("'Rules'!$K$1:$K$5", validations[6].Formula1.Text);
        Assert.Equal(X.DataValidationOperatorValues.GreaterThan, validations[2].Operator.Value);
    }

    [Fact]
    public void Validation_RejectsCommasAndMissingBounds()
    {
        using var p = new ExcelPackage(new MemoryStream());
        var s = p.Workbook.Worksheets.Add("S");
        Assert.Throws<ArgumentException>(() => s.AddDataValidation("A1:A5").List("a,b"));
        Assert.Throws<ArgumentException>(() => s.AddDataValidation("A1:A5").WholeNumber(ValidationOperator.Between, 1));
        Assert.Throws<ArgumentException>(() => s.AddDataValidation("A1:A5").List(Enumerable.Repeat("abcdefghij", 30).ToArray()));
    }

    [Fact]
    public void HyperlinksRichTextCommentsAndCachedFormulas_RoundTrip()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Links");
            var other = (Worksheet?)null;
            s.Row().WriteHyperlink("P4 Software", "https://p4software.com/", tooltip: "Open site")
                   .WriteHyperlink("Mail", "mailto:info@example.com")
                   .WriteInternalLink("Go to data", "'Data'!A1");
            s.Row().WriteRichText("Status: ", new RichText("LATE") { Bold = true, Color = "#C00000", Size = 14 }, new RichText(" (3 days)") { Italic = true });
            s.Row().WriteFormula("1+1", 2.0).WriteFormula("\"a\"&\"b\"", "ab").WriteFormula("1=1", true);
            s.AddComment(1, 1, "Company website", "Glenn");
            s.AddComment(2, 1, "Second note");
            s.AddHyperlink(3, 1, "https://example.com/x?y=1");
            other = wb.Worksheets.Add("Data");
            other.Row().WriteCell("target");
        });

        using (var doc = OpenValidated(bytes))
        {
            var ws = SheetXml(doc, "Links");
            var links = ws.GetFirstChild<X.Hyperlinks>().Elements<X.Hyperlink>().ToList();
            Assert.Equal(4, links.Count);
            var part = ws.WorksheetPart;
            var external = part.HyperlinkRelationships.ToDictionary(r => r.Id, r => r.Uri.ToString());
            Assert.Equal("https://p4software.com/", external[links[0].Id]);
            Assert.Equal("'Data'!A1", links[2].Location.Value);
            Assert.Equal("Open site", links[0].Tooltip.Value);

            var comments = part.WorksheetCommentsPart.Comments;
            Assert.Equal(new[] { "Glenn", "SheetStream" }, comments.Authors.Elements<X.Author>().Select(a => a.Text));
            Assert.Equal(2, comments.CommentList.Elements<X.Comment>().Count());
            Assert.NotNull(part.VmlDrawingParts.Single());

            var styles = doc.WorkbookPart.WorkbookStylesPart.Stylesheet;
            Assert.Contains(styles.CellStyles.Elements<X.CellStyle>(), c => c.Name == "Hyperlink" && c.BuiltinId.Value == 8);
        }

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("Links").ToList();
        Assert.Equal("P4 Software", rows[0][0]);
        Assert.Equal("Status: LATE (3 days)", rows[1][0]);
        Assert.Equal(2.0, rows[2][0]);
        Assert.Equal("ab", rows[2][1]);
        Assert.Equal(true, rows[2][2]);
    }

    [Fact]
    public void TablesCaptureHeadersAndRejectConflicts()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("T");
            var table = s.AddTable("Sales", "A1:C4");
            table.StyleName = "TableStyleLight9";
            table.ShowColumnStripes = true;
            s.Row().WriteCell("Id").WriteCell("Product").WriteCell(2026); // A numeric header is captured as text.
            for (var i = 1; i <= 3; i++) s.Row().WriteCell(i).WriteCell("P" + i).WriteCell(i * 10.5);
            s.AddTable("Explicit", "E1:F3", new[] { "X", "Y" });
        });

        using var doc = OpenValidated(bytes);
        var tables = SheetXml(doc, "T").WorksheetPart.TableDefinitionParts.Select(t => t.Table).OrderBy(t => t.Id.Value).ToList();
        Assert.Equal(2, tables.Count);
        Assert.Equal("A1:C4", tables[0].Reference.Value);
        Assert.Equal(new[] { "Id", "Product", "2026" }, tables[0].TableColumns.Elements<X.TableColumn>().Select(c => c.Name.Value));
        Assert.Equal("TableStyleLight9", tables[0].TableStyleInfo.Name.Value);
        Assert.Equal(new[] { "X", "Y" }, tables[1].TableColumns.Elements<X.TableColumn>().Select(c => c.Name.Value));
    }

    [Theory]
    [InlineData("A1")]
    [InlineData("R1C1")]
    [InlineData("has space")]
    [InlineData("1abc")]
    public void TableNames_AreValidated(string name)
    {
        using var p = new ExcelPackage(new MemoryStream());
        var s = p.Workbook.Worksheets.Add("S");
        Assert.Throws<ArgumentException>(() => s.AddTable(name, "A1:B3"));
    }

    [Fact]
    public void Tables_RejectDuplicatesOverlapsAndLateHeaders()
    {
        using var p = new ExcelPackage(new MemoryStream());
        var s = p.Workbook.Worksheets.Add("S");
        s.AddTable("One", "A1:B3");
        Assert.Throws<ArgumentException>(() => s.AddTable("one", "D1:E3"));
        Assert.Throws<ArgumentException>(() => s.AddTable("Two", "B2:C4"));
        s.Row().WriteCell("a").WriteCell("b");
        s.Row().WriteCell(1);
        Assert.Throws<InvalidOperationException>(() => s.AddTable("Late", "D1:E3"));
    }

    [Fact]
    public void PicturesAndCharts_ProduceValidDrawings()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Charts");
            s.Row().WriteCells("Month", "Sales", "Cost");
            for (var m = 1; m <= 6; m++) s.Row().WriteCell("M" + m).WriteCell(m * 100.0).WriteCell(m * 60.0);
            s.AddPicture(TinyPng, 1, 6, widthPixels: 40);
            foreach (ChartType type in Enum.GetValues(typeof(ChartType)))
            {
                var c = s.AddChart(type, "H" + (2 + (int)type * 16) + ":O" + (15 + (int)type * 16));
                c.Title = type.ToString();
                c.XAxisTitle = "Month";
                c.YAxisTitle = "USD";
                c.DataLabels = type == ChartType.Pie ? ChartDataLabels.Percent : ChartDataLabels.Value;
                c.ValueAxisNumberFormat = "$#,##0";
                c.Legend = LegendPosition.Bottom;
                if (type == ChartType.Scatter || type == ChartType.ScatterLines)
                {
                    c.AddSeries("C2:C7", "B2:B7", "=C1");
                }
                else
                {
                    c.AddSeries("B2:B7", "A2:A7", "=B1", "#4F81BD");
                    c.AddSeries("'Charts'!C2:C7", "A2:A7", "Cost");
                }
            }
        });

        using var doc = OpenValidated(bytes);
        var drawing = SheetXml(doc, "Charts").WorksheetPart.DrawingsPart;
        Assert.Single(drawing.ImageParts);
        Assert.Equal(Enum.GetValues(typeof(ChartType)).Length, drawing.ChartParts.Count());
        var pic = drawing.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Picture>().Single();
        Assert.NotNull(pic);
        var ext = drawing.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.Extent>().Single();
        Assert.Equal(40 * 9525L, ext.Cx.Value);
        Assert.Equal(20 * 9525L, ext.Cy.Value); // Aspect ratio 2:1 kept.
    }

    [Fact]
    public void ImageInspection_ReadsPngJpegAndGifHeaders()
    {
        Assert.Equal(("png", 2, 1), ExcelPicture.Inspect(TinyPng));
        var gif = new byte[] { (byte)'G', (byte)'I', (byte)'F', (byte)'8', (byte)'9', (byte)'a', 10, 0, 20, 0 };
        Assert.Equal(("gif", 10, 20), ExcelPicture.Inspect(gif));
        // SOI, APP0 (length 16), SOF0 with height 300 and width 400.
        var jpeg = new byte[] { 0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x01, 0x2C, 0x01, 0x90, 0, 0 };
        Assert.Equal(("jpeg", 400, 300), ExcelPicture.Inspect(jpeg));
        Assert.Throws<ArgumentException>(() => ExcelPicture.Inspect(new byte[] { 1, 2, 3, 4 }));
    }

    [Fact]
    public void ProtectionPageSetupNamesAndSheetSettings_AreWritten()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            wb.Properties.Title = "Report";
            wb.Properties.Author = "P4 Software";
            wb.Properties.Company = "P4 Software";
            wb.Protect("secret");
            var unlocked = wb.Styles.Create(o => { o.Protection.Locked = false; o.Indent = 2; o.TextRotation = 45; o.ShrinkToFit = true; });
            var hidden = wb.Worksheets.Add("Hidden");
            hidden.Visibility = SheetVisibility.Hidden;
            var s = wb.Worksheets.Add("Main");
            s.TabColor = "#C00000";
            s.ShowGridLines = false;
            s.Zoom = 85;
            s.PageSetup.FitToWidth = 1;
            s.PageSetup.Orientation = PageOrientation.Landscape;
            s.PageSetup.PaperSize = PaperSize.A4;
            s.PageSetup.Header = "&LP4&RConfidential";
            s.PageSetup.Footer = "Page &P of &N";
            s.PageSetup.PrintTitleRows = (1, 1);
            s.PageSetup.PrintArea = "A1:C10";
            s.PageSetup.PrintGridlines = true;
            s.PageSetup.Margins = (0.5, 0.5, 0.6, 0.6, 0.2, 0.2);
            s.Column(2).SetOutlineLevel(1);
            s.Protect("pw", o => { o.AllowSort = true; o.AllowFormatColumns = true; o.AllowSelectLockedCells = false; });
            s.Row().WriteCell("Input", unlocked);
            s.Row(2, new RowOptions { OutlineLevel = 1, Hidden = true, Height = 20 }).WriteCell(1);
            wb.DefinedNames.Add("TaxRate", "'Main'!B2");
            wb.DefinedNames.Add("Local", "A1:A3", scope: s);
            wb.DefinedNames.Add("Constant", "=0.07");
        });

        using var doc = OpenValidated(bytes);
        var book = doc.WorkbookPart.Workbook;
        Assert.Equal(X.SheetStateValues.Hidden, book.Sheets.Elements<X.Sheet>().First().State.Value);
        Assert.Equal(1u, book.BookViews.Elements<X.WorkbookView>().Single().ActiveTab.Value);
        Assert.Equal("SHA-512", book.WorkbookProtection.WorkbookAlgorithmName.Value);
        var names = book.DefinedNames.Elements<X.DefinedName>().ToDictionary(d => d.Name.Value + (d.LocalSheetId == null ? "" : "@" + d.LocalSheetId.Value), d => d.Text);
        Assert.Equal("'Main'!$A$1:$C$10", names["_xlnm.Print_Area@1"]);
        Assert.Equal("'Main'!$1:$1", names["_xlnm.Print_Titles@1"]);
        Assert.Equal("'Main'!$B$2", names["TaxRate"]);
        Assert.Equal("'Main'!$A$1:$A$3", names["Local@1"]);
        Assert.Equal("0.07", names["Constant"]);
        Assert.Equal("Report", doc.PackageProperties.Title);
        Assert.Equal("P4 Software", doc.PackageProperties.Creator);

        var ws = SheetXml(doc, "Main");
        Assert.Equal("FFC00000", ws.SheetProperties.TabColor.Rgb.Value);
        Assert.True(ws.SheetProperties.PageSetupProperties.FitToPage.Value);
        Assert.False(ws.SheetViews.Elements<X.SheetView>().Single().ShowGridLines.Value);
        Assert.Equal(X.OrientationValues.Landscape, ws.GetFirstChild<X.PageSetup>().Orientation.Value);
        Assert.Equal("&LP4&RConfidential", ws.GetFirstChild<X.HeaderFooter>().OddHeader.Text);
        Assert.Equal("&CPage &P of &N", ws.GetFirstChild<X.HeaderFooter>().OddFooter.Text);
        var protection = ws.GetFirstChild<X.SheetProtection>();
        Assert.True(protection.Sheet.Value);
        Assert.False(protection.Sort.Value);           // "0" = sorting allowed.
        Assert.True(protection.SelectLockedCells.Value); // "1" = selecting locked cells blocked.
        var row2 = ws.Descendants<X.Row>().Single(r => r.RowIndex.Value == 2);
        Assert.Equal(1, row2.OutlineLevel.Value);
        Assert.True(row2.Hidden.Value);
        var unlockedXf = doc.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.Elements<X.CellFormat>()
            .ElementAt((int)ws.Descendants<X.Cell>().First().StyleIndex.Value);
        Assert.False(unlockedXf.Protection.Locked.Value);
        Assert.Equal(2u, unlockedXf.Alignment.Indent.Value);
        Assert.Equal(45u, unlockedXf.Alignment.TextRotation.Value);
    }

    [Fact]
    public void PasswordHasher_MatchesKnownVerifier()
    {
        // Independent check of the ISO 29500 algorithm: hash(salt+pwd), then 3 rounds of hash(prev + LE counter).
        var salt = Enumerable.Range(0, 16).Select(i => (byte)i).ToArray();
        using var sha = System.Security.Cryptography.SHA512.Create();
        var h = sha.ComputeHash(salt.Concat(System.Text.Encoding.Unicode.GetBytes("pw")).ToArray());
        for (var i = 0; i < 3; i++) h = sha.ComputeHash(h.Concat(BitConverter.GetBytes(i)).ToArray());
        Assert.Equal(h, Internal.PasswordHasher.Compute("pw", salt, 3));
    }

    [Fact]
    public void AllSheetsHidden_IsRejected()
    {
        using var p = new ExcelPackage(new MemoryStream());
        p.Workbook.Worksheets.Add("Only").Visibility = SheetVisibility.Hidden;
        Assert.Throws<InvalidOperationException>(() => p.Save());
    }

    [Fact]
    public void EmptyWorkbook_GetsADefaultSheet()
    {
        var bytes = ExcelPackage.CreateBytes(_ => { });
        using var doc = OpenValidated(bytes);
        Assert.Equal("Sheet1", doc.WorkbookPart.Workbook.Sheets.Elements<X.Sheet>().Single().Name.Value);
    }

    [Fact]
    public void AutoFit_SizesColumnsFromContentAndAllowsLateColumnSettings()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var bold = wb.Styles.Create(o => { o.Font.Bold = true; o.Font.Size = 16; });
            var s = wb.Worksheets.Add("Fit");
            s.AutoFitColumns(minWidth: 4, maxWidth: 60);
            s.Row().WriteCell("Id").WriteCell("A fairly long product description here").WriteCell("Big", bold).WriteCell(new string('x', 500));
            s.Row().WriteCell(123456789).WriteCell("short").WriteCell(1.5).WriteCell(DateTime.Today);
            s.Column(1).SetWidth(20); // Still allowed: the header is deferred in AutoFit mode.
            s.FreezePanes(1);
        });

        using var doc = OpenValidated(bytes);
        var cols = SheetXml(doc, "Fit").GetFirstChild<X.Columns>().Elements<X.Column>().ToDictionary(c => c.Min.Value, c => c.Width.Value);
        Assert.Equal(20, cols[1]);                       // Explicit width wins.
        Assert.InRange(cols[2], 38, 45);                 // ~38 characters + padding.
        Assert.InRange(cols[3], 6, 12);                  // "Big" at 16pt bold is scaled up.
        Assert.Equal(60, cols[4]);                       // Clamped to maxWidth.
        Assert.Equal(2, SheetXml(doc, "Fit").Descendants<X.Row>().Count());
    }

    [Fact]
    public void RandomAccessCells_WriteInAnyOrder()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Grid");
            var money = wb.Styles.Create(o => o.NumberFormat = "#,##0.00");
            s.Cells[3, 2].Value = 30;
            s.Cells["A1"].Value = "Header";
            s.Cells[2, 3].Value = 2.5;
            s.Cells[2, 3].Style = money;
            s.Cells[4, 2].Formula = "SUM(B3:B3)";
            s.Cells[2, 1].Value = new DateTime(2026, 10, 13);
            Assert.Equal(30, s.Cells[3, 2].Value);
            s.Column(2).SetWidth(30); // Allowed any time in random-access mode.
            Assert.Throws<InvalidOperationException>(() => s.Row());
        });

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var rows = reader.ReadRows("Grid").ToList();
        Assert.Equal(new[] { 1, 2, 3, 4 }, rows.Select(r => r.Index));
        Assert.Equal("Header", rows[0][0]);
        Assert.Equal(new DateTime(2026, 10, 13), rows[1][0]);
        Assert.Equal(2.5, rows[1][2]);
        Assert.Equal(30.0, rows[2][1]);
        using var doc = OpenValidated(bytes);
    }

    [Fact]
    public void StreamingSheet_RejectsRandomAccess()
    {
        using var p = new ExcelPackage(new MemoryStream());
        var s = p.Workbook.Worksheets.Add("S");
        s.Row().WriteCell(1);
        Assert.Throws<InvalidOperationException>(() => s.Cells[1, 1].Value = 2);
    }

    public sealed class Order
    {
        [SheetColumn("Order #", Order = 0)]
        public int Id { get; set; }

        [DisplayName("Customer name")]
        public string? Customer { get; set; }

        [SheetColumn(NumberFormat = "#,##0.00")]
        public decimal Amount { get; set; }

        public DateTime Placed { get; set; }

        public bool Paid { get; set; }

        public Status State { get; set; }

        [SheetColumn(Ignore = true)]
        public string Secret { get; set; } = "hidden";

        public double? Discount { get; set; }
    }

    public enum Status { Draft, Sent, Closed }

    [Fact]
    public void LoadFromCollection_WritesTableAndReaderMapsItBack()
    {
        var orders = Enumerable.Range(1, 25).Select(i => new Order
        {
            Id = i, Customer = "C" + i, Amount = i * 10.25m, Placed = new DateTime(2026, 1, i % 28 + 1),
            Paid = i % 2 == 0, State = (Status)(i % 3), Discount = i % 4 == 0 ? null : i / 100.0,
        }).ToList();

        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var header = wb.Styles.Create(o => o.Font.Bold = true);
            var s = wb.Worksheets.Add("Orders");
            s.Row().WriteCell("Report title");
            var result = s.LoadFromCollection(orders, new LoadOptions { HeaderStyle = header, TableName = "Orders", StartColumn = 2 });
            Assert.Equal((2, 27, 2, 8, 25), (result.FirstRow, result.LastRow, result.FirstColumn, result.LastColumn, result.DataRows));
            var dyn = wb.Worksheets.Add("Dynamic");
            dyn.LoadFromCollection(new[]
            {
                new Dictionary<string, object?> { ["a"] = 1, ["b"] = "x" },
                new Dictionary<string, object?> { ["a"] = 2, ["b"] = null },
            }, new LoadOptions { AutoFilter = true });
        });

        using (var doc = OpenValidated(bytes))
        {
            var table = SheetXml(doc, "Orders").WorksheetPart.TableDefinitionParts.Single().Table;
            Assert.Equal("B2:H27", table.Reference.Value);
            Assert.Equal(new[] { "Order #", "Customer name", "Amount", "Placed", "Paid", "State", "Discount" },
                table.TableColumns.Elements<X.TableColumn>().Select(c => c.Name.Value));
        }

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        var dt = reader.ReadDataTable("Dynamic");
        Assert.Equal(new[] { "a", "b" }, dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
        Assert.Equal(2, dt.Rows.Count);

        // Read<T> needs the header in row 1, so read the table area through a copy without the title row.
        var copy = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("O").LoadFromCollection(orders));
        using var r2 = ExcelReader.Open(new MemoryStream(copy));
        var back = r2.Read<Order>("O").ToList();
        Assert.Equal(25, back.Count);
        for (var i = 0; i < 25; i++)
        {
            Assert.Equal(orders[i].Id, back[i].Id);
            Assert.Equal(orders[i].Customer, back[i].Customer);
            Assert.Equal(orders[i].Amount, back[i].Amount);
            Assert.Equal(orders[i].Placed, back[i].Placed);
            Assert.Equal(orders[i].Paid, back[i].Paid);
            Assert.Equal(orders[i].State, back[i].State);
            Assert.Equal(orders[i].Discount, back[i].Discount);
            Assert.Equal("hidden", back[i].Secret);
        }
    }

    [Fact]
    public void LoadFromDataTableAndReader_StreamRows()
    {
        var table = new DataTable();
        table.Columns.Add("Name", typeof(string));
        table.Columns.Add("Qty", typeof(int)).Caption = "Quantity";
        table.Columns.Add("When", typeof(DateTime));
        table.Rows.Add("a", 1, new DateTime(2026, 1, 1));
        table.Rows.Add("b", DBNull.Value, new DateTime(2026, 1, 2));

        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            wb.Worksheets.Add("FromTable").LoadFromDataTable(table, new LoadOptions { TableName = "Tbl" });
            wb.Worksheets.Add("FromReader").LoadFromDataReader(table.CreateDataReader());
        });

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        foreach (var sheet in new[] { "FromTable", "FromReader" })
        {
            var rows = reader.ReadRows(sheet).ToList();
            Assert.Equal(3, rows.Count);
            Assert.Equal(sheet == "FromTable" ? "Quantity" : "Qty", rows[0][1]);
            Assert.Equal(1.0, rows[1][1]);
            Assert.Null(rows[2][1]);
            Assert.Equal(new DateTime(2026, 1, 2), rows[2][2]);
        }
        using var doc = OpenValidated(bytes);
    }

    [Fact]
    public void Reader_HandlesTypesErrorsAndSparseCells()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Types");
            var pct = wb.Styles.Create(o => o.NumberFormat = "0.00%");
            var custDate = wb.Styles.Create(o => o.NumberFormat = "dd \"de\" mmmm yyyy");
            var colorFmt = wb.Styles.Create(o => o.NumberFormat = "[Red]#,##0");
            s.Row().WriteCell("text").WriteCell(1.25).WriteCell(true).WriteCell(double.NaN);
            s.Row(3).Skip(2).WriteCell(0.5, pct).WriteCell(new DateTime(2026, 3, 4), custDate).WriteCell(TimeSpan.FromHours(30)).WriteCell(7.0, colorFmt);
            s.Row().WriteInlineString("_x0041_ literal\u0001ctrl");
        });

        using var reader = ExcelReader.Open(new MemoryStream(bytes));
        Assert.Equal(new[] { "Types" }, reader.SheetNames);
        var rows = reader.ReadRows(0).ToList();
        Assert.Equal(new[] { 1, 3, 4 }, rows.Select(r => r.Index));
        Assert.Equal("text", rows[0][0]);
        Assert.Equal(1.25, rows[0].GetDouble(1));
        Assert.Equal(true, rows[0][2]);
        Assert.Equal("#NUM!", ((ExcelErrorValue)rows[0][3]!).Code);
        Assert.Null(rows[1][0]);
        Assert.Equal(0.5, rows[1][2]);                                    // Percent is a number, not a date.
        Assert.Equal(new DateTime(2026, 3, 4), rows[1][3]);              // Custom date code with quoted text.
        Assert.True(Math.Abs(((TimeSpan)rows[1][4]!).TotalHours - 30) < 0.001);
        Assert.Equal(7.0, rows[1][5]);                                    // [Red] is a colour, not elapsed time.
        Assert.Equal("_x0041_ literal\u0001ctrl", rows[2][0]);            // Escapes reversed exactly.
        Assert.Throws<ArgumentException>(() => reader.ReadRows("Nope").ToList());
    }

    [Fact]
    public void Reader_RejectsHostileColumnReferencesAndPlacesOutOfOrderCells()
    {
        // Build a valid workbook, then rewrite its sheet XML by hand to simulate a crafted file.
        var bytes = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S").Row().WriteCell(1.0).WriteCell(2.0));
        byte[] Rewrite(string sheetXml)
        {
            var ms = new MemoryStream();
            ms.Write(bytes, 0, bytes.Length);
            using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
            {
                zip.GetEntry("xl/worksheets/sheet1.xml")!.Delete();
                using var w = new StreamWriter(zip.CreateEntry("xl/worksheets/sheet1.xml").Open());
                w.Write(sheetXml);
            }
            return ms.ToArray();
        }
        const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        var hostile = Rewrite($"<worksheet xmlns=\"{ns}\"><sheetData><row r=\"1\"><c r=\"ZZZZZZ1\"><v>1</v></c></row></sheetData></worksheet>");
        using (var r = ExcelReader.Open(new MemoryStream(hostile)))
            Assert.Throws<InvalidDataException>(() => r.ReadRows(0).ToList());

        var outOfOrder = Rewrite($"<worksheet xmlns=\"{ns}\"><sheetData><row r=\"1\"><c r=\"C1\"><v>3</v></c><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>");
        using (var r = ExcelReader.Open(new MemoryStream(outOfOrder)))
        {
            var row = r.ReadRows(0).Single();
            Assert.Equal(1.0, row[0]);
            Assert.Null(row[1]);
            Assert.Equal(3.0, row[2]);
        }
    }

    [Fact]
    public void Reader_RejectsNonZipInput()
    {
        var ex = Assert.Throws<InvalidDataException>(() => ExcelReader.Open(new MemoryStream(new byte[] { 1, 2, 3, 4, 5 })));
        Assert.Contains(".xlsx", ex.Message);
    }

    [Fact]
    public void FullFeatureWorkbook_StillStreamsLargeData()
    {
        // Tables, conditional formats and validations must not change the flat-memory profile of row writing.
        const int rows = 100_000;
        var before = GC.GetTotalMemory(true);
        using var sink = new MemoryStream();
        using (var p = new ExcelPackage(sink, leaveOpen: true, new PackageOptions { CompressionLevel = System.IO.Compression.CompressionLevel.Fastest }))
        {
            var s = p.Workbook.Worksheets.Add("Big");
            var t = s.AddTable("Big", 1, 1, rows + 1, 3);
            s.AddConditionalFormatting(2, 3, rows + 1, 3).DataBar("#638EC6");
            s.AddDataValidation(2, 2, rows + 1, 2).List("x", "y");
            s.Row().WriteCells("A", "B", "C");
            for (var i = 0; i < rows; i++) s.Row().WriteCell(i).WriteCell(i % 2 == 0 ? "x" : "y").WriteCell(i * 0.5);
            p.Save();
            var retained = GC.GetTotalMemory(true) - before - sink.Length;
            Assert.True(retained < 32 * 1024 * 1024, $"Retained {retained / 1024} KB beyond the output itself.");
        }
        Assert.True(sink.Length > 500_000);
    }
}
