#pragma warning disable CS8600, CS8602, CS8603, CS8604, CS8619, CS8714, CS8621
using System.IO.Compression;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SheetStream.Reading;
using SheetStream.Styles;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace SheetStream.Tests;

/// <summary>Editing a small existing workbook: open, read and change cells, save, with everything else preserved.</summary>
public class EditTests
{
    /// <summary>
    /// A workbook as a user might have it: a tall title row merged across the sheet, a header, three data rows
    /// (text, a date, money, a link), a SUM formula, a boolean and an error, plus a second sheet with a table, a
    /// chart and a picture that must survive untouched.
    /// </summary>
    private static byte[] BuildWorkbook()
    {
        return ExcelPackage.CreateBytes(wb =>
        {
            var title = wb.Styles.Create(o => { o.Font.Bold = true; o.Font.Size = 16; });
            var header = wb.Styles.Create(o => { o.Font.Bold = true; o.Fill.BackgroundColor = "#DDEBF7"; });
            var money = wb.Styles.Create(o => o.NumberFormat = "$#,##0.00");
            var s = wb.Worksheets.Add("Data");
            s.Column(2).SetWidth(14);
            s.FreezePanes(2);
            s.MergeCells("A1:D1");
            s.Row(1, title, 30).WriteCell("Sales report");
            s.Row(2).WriteCell("Sku", header).WriteCell("Date", header).WriteCell("Amount", header).WriteCell("Link", header);
            s.Row(3).WriteCell("SKU-1").WriteCell(new DateTime(2026, 10, 1)).WriteCell(10.5, money).WriteHyperlink("docs", "https://example.com/1");
            s.Row(4).WriteCell("SKU-2").WriteCell(new DateTime(2026, 10, 2)).WriteCell(20.25, money).WriteInlineString("  padded  ");
            s.Row(5).WriteCell("SKU-3").WriteCell(new DateTime(2026, 10, 3)).WriteCell(30.0, money).WriteCell(true);
            s.Row(6).WriteCell("Total", header).WriteEmpty(header).WriteFormula("SUM(C3:C5)", 60.75, money).WriteCell(double.NaN);

            var second = wb.Worksheets.Add("Chart");
            second.AddTable("Figures", "A1:B4");
            second.Row(1).WriteCell("Month").WriteCell("Value");
            second.Row(2).WriteCell("Jan").WriteCell(1);
            second.Row(3).WriteCell("Feb").WriteCell(2);
            second.Row(4).WriteCell("Mar").WriteCell(3);
            var chart = second.AddChart(ChartType.Column, "D2:J16");
            chart.AddSeries(values: "B2:B4", categories: "A2:A4", name: "=B1");
            second.AddPicture(FeatureTests.TinyPng, 18, 1);
        });
    }

    private static SpreadsheetDocument OpenValidated(byte[] bytes)
    {
        var doc = SpreadsheetDocument.Open(new MemoryStream(bytes), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2016).Validate(doc).ToList();
        Assert.True(errors.Count == 0, string.Join(Environment.NewLine, errors.Select(e => $"{e.Path?.PartUri} {e.Path?.XPath}: {e.Description}")));
        return doc;
    }

    private static byte[] EditBytes(byte[] input, Action<ExcelPackage> edit, PackageOptions? options = null)
    {
        using var output = new MemoryStream();
        using (var package = ExcelPackage.Edit(new MemoryStream(input), output, leaveOpen: true, options))
        {
            edit(package);
            package.Save();
        }
        return output.ToArray();
    }

    private static Dictionary<string, X.Cell> CellsOf(SpreadsheetDocument doc, string sheetName)
    {
        var book = doc.WorkbookPart;
        var sheet = book.Workbook.Sheets.Elements<X.Sheet>().Single(x => x.Name.Value == sheetName);
        return ((WorksheetPart)book.GetPartById(sheet.Id)).Worksheet.Descendants<X.Cell>().ToDictionary(c => c.CellReference.Value);
    }

    [Fact]
    public void Edit_ReadsChangesAndPreservesEverythingElse()
    {
        var original = BuildWorkbook();
        var edited = EditBytes(original, p =>
        {
            var s = p.Workbook.Worksheets.Edit("data"); // Case-insensitive.

            // Existing values surface typed, like ExcelReader returns them.
            Assert.Equal("Sales report", s.Cells["A1"].Value);
            Assert.Equal("Sku", s.Cells[2, 1].Value);
            Assert.Equal(new DateTime(2026, 10, 2), s.Cells["B4"].Value);
            Assert.Equal(20.25, s.Cells["C4"].Value);
            Assert.Equal("  padded  ", s.Cells["D4"].Value);
            Assert.Equal(true, s.Cells["D5"].Value);
            Assert.Equal("SUM(C3:C5)", s.Cells["C6"].Formula);
            Assert.Equal(60.75, s.Cells["C6"].Value);                 // Cached result.
            Assert.Equal("#NUM!", ((ExcelErrorValue)s.Cells["D6"].Value).Code);
            Assert.NotNull(s.Cells["C3"].Style);
            Assert.Null(s.Cells["A3"].Style);
            Assert.Equal((6, 4), s.Cells.Dimension);
            Assert.Equal(21, s.Cells.Existing.Count());
            Assert.Null(s.Cells["Z99"].Value);

            // Change a cell, keep the money format on it; clear one; empty one but keep its style; append a row.
            s.Cells["C4"].Value = 99.5;
            s.Cells["D4"].Clear();
            s.Cells["A5"].Value = null;
            s.Cells["A5"].Style = s.TemplateStyle("A2");
            var next = s.Cells.Dimension.LastRow + 1;
            s.Cells[next, 1].Value = "SKU-NEW";
            s.Cells[next, 2].Value = new DateTime(2026, 10, 4);
            s.Cells[next, 3].Value = 5;
            s.Cells[next, 3].Style = s.Cells["C3"].Style;
            Assert.Equal(99.5, s.Cells["C4"].Value);
            Assert.Throws<InvalidOperationException>(() => s.Row(8)); // Edit mode is random access only.
        });

        using var doc = OpenValidated(edited);
        var book = doc.WorkbookPart;
        Assert.Equal(new[] { "Data", "Chart" }, book.Workbook.Sheets.Elements<X.Sheet>().Select(x => x.Name.Value));
        Assert.True(book.Workbook.CalculationProperties.FullCalculationOnLoad.Value);

        var dataPart = (WorksheetPart)book.GetPartById(book.Workbook.Sheets.Elements<X.Sheet>().First().Id);
        var ws = dataPart.Worksheet;
        var cells = CellsOf(doc, "Data");
        var originalCells = CellsOf(OpenValidated(original), "Data");

        Assert.Equal("99.5", cells["C4"].CellValue.Text);
        Assert.Equal(originalCells["C4"].StyleIndex.Value, cells["C4"].StyleIndex.Value);        // Money format kept.
        Assert.Equal(originalCells["C3"].OuterXml, cells["C3"].OuterXml);                        // Untouched cells are byte-identical.
        Assert.Equal(originalCells["B3"].OuterXml, cells["B3"].OuterXml);
        Assert.Equal(originalCells["D4"].OuterXml, originalCells["D4"].OuterXml);
        Assert.False(cells.ContainsKey("D4"));                                                   // Cleared.
        Assert.Null(cells["A5"].CellValue);                                                      // Emptied, styled.
        Assert.Equal(originalCells["A2"].StyleIndex.Value, cells["A5"].StyleIndex.Value);
        Assert.Equal("SUM(C3:C5)", cells["C6"].CellFormula.Text);
        Assert.Equal("60.75", cells["C6"].CellValue.Text);                                       // Cached value survives.
        Assert.Equal(X.CellValues.Error, cells["D6"].DataType.Value);
        Assert.Equal(X.CellValues.SharedString, cells["A7"].DataType.Value);
        Assert.Equal("5", cells["C7"].CellValue.Text);
        Assert.Equal(originalCells["C3"].StyleIndex.Value, cells["C7"].StyleIndex.Value);

        // Layout around the data: merge, hyperlink, frozen pane, column width, row height, no stale dimension.
        Assert.Equal("A1:D1", ws.GetFirstChild<X.MergeCells>().Elements<X.MergeCell>().Single().Reference.Value);
        Assert.Single(ws.GetFirstChild<X.Hyperlinks>().Elements<X.Hyperlink>());
        Assert.Single(dataPart.HyperlinkRelationships);
        Assert.NotNull(ws.GetFirstChild<X.SheetViews>().Descendants<X.Pane>().Single());
        Assert.Equal(14, ws.GetFirstChild<X.Columns>().Elements<X.Column>().Single().Width.Value);
        Assert.Null(ws.SheetDimension);
        var rows = ws.Descendants<X.Row>().ToList();
        Assert.Equal(new uint[] { 1, 2, 3, 4, 5, 6, 7 }, rows.Select(r => r.RowIndex.Value));
        Assert.Equal(30, rows[0].Height.Value);
        Assert.True(rows[0].CustomHeight.Value);

        // The untouched sheet keeps its table, chart and picture.
        var chartPart = (WorksheetPart)book.GetPartById(book.Workbook.Sheets.Elements<X.Sheet>().Last().Id);
        Assert.Equal("A1:B4", chartPart.TableDefinitionParts.Single().Table.Reference.Value);
        Assert.Single(chartPart.DrawingsPart.ChartParts);
        Assert.Single(chartPart.DrawingsPart.ImageParts);

        // The reader sees the same picture.
        using var reader = ExcelReader.Open(new MemoryStream(edited));
        var data = reader.ReadRows("Data").ToDictionary(r => r.Index);
        Assert.Equal("Sales report", data[1][0]);
        Assert.Equal(new DateTime(2026, 10, 1), data[3][1]);
        Assert.Equal(99.5, data[4][2]);
        Assert.Equal("SKU-NEW", data[7][0]);
        Assert.Equal(new DateTime(2026, 10, 4), data[7][1]);
        Assert.Equal("Feb", reader.ReadRows("Chart").Single(r => r.Index == 3)[0]);
    }

    [Fact]
    public void Edit_AppendedRowsExtendATableEndingOnTheLastRow()
    {
        var edited = EditBytes(BuildWorkbook(), p =>
        {
            var s = p.Workbook.Worksheets.Edit(1);           // "Chart", by position.
            Assert.Equal((4, 2), s.Cells.Dimension);
            s.Cells[5, 1].Value = "Apr";
            s.Cells[5, 2].Value = 4;
        });
        using var doc = OpenValidated(edited);
        var book = doc.WorkbookPart;
        var chartPart = (WorksheetPart)book.GetPartById(book.Workbook.Sheets.Elements<X.Sheet>().Last().Id);
        Assert.Equal("A1:B5", chartPart.TableDefinitionParts.Single().Table.Reference.Value);
        Assert.Equal("A1:B5", chartPart.TableDefinitionParts.Single().Table.AutoFilter.Reference.Value);
    }

    [Fact]
    public void Edit_CanMixEditedAndStreamedSheets()
    {
        var edited = EditBytes(BuildWorkbook(), p =>
        {
            p.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "Edited title";
            p.Workbook.Worksheets.Open("Chart").Row(6).WriteCell("streamed");  // Sealing Data, streaming into Chart.
        });
        using var doc = OpenValidated(edited);
        using var reader = ExcelReader.Open(new MemoryStream(edited));
        Assert.Equal("Edited title", reader.ReadRows("Data").First()[0]);
        var chart = reader.ReadRows("Chart").ToDictionary(r => r.Index);
        Assert.Equal("Mar", chart[4][0]);
        Assert.Equal("streamed", chart[6][0]);
    }

    [Fact]
    public void Edit_RoundTripsAnUntouchedSheetUnchanged()
    {
        // Opening a sheet for editing and saving without changes must reproduce every cell.
        var original = BuildWorkbook();
        var edited = EditBytes(original, p => p.Workbook.Worksheets.Edit("Data"));
        using var before = OpenValidated(original);
        using var after = OpenValidated(edited);
        var a = CellsOf(before, "Data");
        var b = CellsOf(after, "Data");
        Assert.Equal(a.Keys.OrderBy(k => k), b.Keys.OrderBy(k => k));
        foreach (var key in a.Keys) Assert.Equal(a[key].OuterXml, b[key].OuterXml);
    }

    [Fact]
    public void Edit_InPlace_ReplacesTheOriginalAndLeavesNoTempFile()
    {
        var dir = Path.Combine(Path.GetTempPath(), "ss-edit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "book.xlsx");
        try
        {
            File.WriteAllBytes(path, BuildWorkbook());
            using (var package = ExcelPackage.Edit(path))
            {
                var s = package.Workbook.Worksheets.Edit("Data");
                s.Cells["A1"].Value = "In place";
            } // Disposing saves and swaps the file in.

            Assert.Single(Directory.GetFiles(dir));
            using (var reader = ExcelReader.Open(path))
                Assert.Equal("In place", reader.ReadRows("Data").First()[0]);

            // Edit(input, output) with the same path means the same thing.
            using (var package = ExcelPackage.Edit(path, path))
            {
                package.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "Twice";
                package.Save();
            }
            Assert.Single(Directory.GetFiles(dir));
            using (var reader = ExcelReader.Open(path))
                Assert.Equal("Twice", reader.ReadRows("Data").First()[0]);
            using var doc = OpenValidated(File.ReadAllBytes(path));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Edit_InPlace_FailedSaveLeavesTheOriginalIntact()
    {
        var dir = Path.Combine(Path.GetTempPath(), "ss-edit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "book.xlsx");
        try
        {
            var original = BuildWorkbook();
            File.WriteAllBytes(path, original);
            var ex = Record.Exception(() =>
            {
                using var package = ExcelPackage.Edit(path);
                var s = package.Workbook.Worksheets.Edit("Data");
                s.Cells["A1"].Value = "Lost";
                s.MergeCells("A9:B9"); // Layout comes from the file; this makes the save throw.
                package.Save();
            });
            Assert.IsType<NotSupportedException>(ex);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(dir));          // The temporary file is gone.
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Edit_FileToFile_WritesACopy()
    {
        var dir = Path.Combine(Path.GetTempPath(), "ss-edit-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var input = Path.Combine(dir, "in.xlsx");
            var output = Path.Combine(dir, "out.xlsx");
            var original = BuildWorkbook();
            File.WriteAllBytes(input, original);
            using (var package = ExcelPackage.Edit(input, output))
                package.Workbook.Worksheets.Edit("Data").Cells["A1"].Value = "Copy";
            Assert.Equal(original, File.ReadAllBytes(input));
            using var reader = ExcelReader.Open(output);
            Assert.Equal("Copy", reader.ReadRows("Data").First()[0]);
            Assert.Throws<FileNotFoundException>(() => ExcelPackage.Edit(Path.Combine(dir, "missing.xlsx")));
        }
        finally
        {
            Directory.Delete(dir, recursive: true);
        }
    }

    [Fact]
    public void Edit_RefusesSheetsOverTheCellLimit()
    {
        var big = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("Big");
            for (var r = 1; r <= 50; r++) s.Row().WriteCells(r, r, r, r); // 200 cells
            wb.Worksheets.Add("Small").Row().WriteCell("ok");
        });
        using var output = new MemoryStream();
        using var package = ExcelPackage.Edit(new MemoryStream(big), output, leaveOpen: true, new PackageOptions { MaxEditableCells = 199 });
        var ex = Assert.Throws<InvalidOperationException>(() => package.Workbook.Worksheets.Edit("Big"));
        Assert.Contains("small workbooks", ex.Message);
        Assert.Contains("MaxEditableCells", ex.Message);
        // The refusal left the package usable: the small sheet can still be edited and the big one copied.
        package.Workbook.Worksheets.Edit("Small").Cells["A1"].Value = "edited";
        package.Save();
        using var reader = ExcelReader.Open(new MemoryStream(output.ToArray()));
        Assert.Equal(50, reader.ReadRows("Big").Count());
        Assert.Equal("edited", reader.ReadRows("Small").First()[0]);
    }

    /// <summary>Rewrites one sheet part of a workbook, for files SheetStream itself does not produce.</summary>
    private static byte[] PatchSheet(byte[] workbook, string part, Func<string, string> patch)
    {
        var ms = new MemoryStream();
        ms.Write(workbook, 0, workbook.Length);
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = zip.GetEntry(part)!;
            string xml;
            using (var sr = new StreamReader(entry.Open())) xml = sr.ReadToEnd();
            entry.Delete();
            using var w = new StreamWriter(zip.CreateEntry(part).Open());
            w.Write(patch(xml));
        }
        return ms.ToArray();
    }

    [Fact]
    public void Edit_SharedFormulas_RoundTripAndRefuseEditingTheMaster()
    {
        // Excel writes a formula filled down as one master with the text and dependents that point at it.
        var plain = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell(1).WriteCell(2);
            s.Row(2).WriteCell(3).WriteFormula("A2*2", 6.0);
            s.Row(3).WriteCell(4).WriteCell(0);
        });
        var shared = PatchSheet(plain, "xl/worksheets/sheet1.xml", xml =>
        {
            Assert.Contains("<c r=\"B2\"><f>A2*2</f>", xml);
            Assert.Contains("<c r=\"B3\"><v>0</v></c>", xml);
            return xml.Replace("<c r=\"B2\"><f>A2*2</f>", "<c r=\"B2\"><f t=\"shared\" ref=\"B2:B3\" si=\"0\">A2*2</f>")
                      .Replace("<c r=\"B3\"><v>0</v></c>", "<c r=\"B3\"><f t=\"shared\" si=\"0\"/><v>8</v></c>");
        });

        var edited = EditBytes(shared, p =>
        {
            var s = p.Workbook.Worksheets.Edit("S");
            Assert.Equal("A2*2", s.Cells["B2"].Formula);
            Assert.Null(s.Cells["B3"].Formula);           // A dependent has no text of its own.
            Assert.Equal(8.0, s.Cells["B3"].Value);
            var ex = Assert.Throws<NotSupportedException>(() => s.Cells["B2"].Value = 1);
            Assert.Contains("B2", ex.Message);
            Assert.Throws<NotSupportedException>(() => s.Cells["B2"].Clear());
            s.Cells["B2"].Style = p.Workbook.Styles.Create(o => o.Font.Bold = true); // Style changes are fine.
            s.Cells["A3"].Value = 40;                                                 // So is editing elsewhere.
        });

        using var doc = OpenValidated(edited);
        var cells = CellsOf(doc, "S");
        Assert.Equal(X.CellFormulaValues.Shared, cells["B2"].CellFormula.FormulaType.Value);
        Assert.Equal("B2:B3", cells["B2"].CellFormula.Reference.Value);
        Assert.Equal("A2*2", cells["B2"].CellFormula.Text);
        Assert.NotNull(cells["B2"].StyleIndex);
        Assert.Equal(X.CellFormulaValues.Shared, cells["B3"].CellFormula.FormulaType.Value);
        Assert.Equal(0u, cells["B3"].CellFormula.SharedIndex.Value);
        Assert.Equal("8", cells["B3"].CellValue.Text);
        Assert.Equal("40", cells["A3"].CellValue.Text);
    }

    [Fact]
    public void Edit_PreservesRowAttributesRichTextAndNamespacedMarkup()
    {
        // Rows hidden or with an outline level, an empty row carrying only a height, rich inline text and an
        // x14ac:dyDescent attribute like Excel writes: none of it is modelled, all of it must survive.
        var plain = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1, new RowOptions { Hidden = true }).WriteCell("hidden");
            s.Row(2, new RowOptions { OutlineLevel = 1 }).WriteRichText("Keep ", new RichText("bold") { Bold = true });
            s.Row(4, null, 42).WriteCell("tall");
        });
        var patched = PatchSheet(plain, "xl/worksheets/sheet1.xml", xml =>
        {
            Assert.Contains("<row r=\"4\" ht=\"42\" customHeight=\"1\">", xml);
            return xml.Replace("<worksheet ", "<worksheet xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x14ac\" ")
                      .Replace("<row r=\"4\" ht=\"42\" customHeight=\"1\">", "<row r=\"3\" ht=\"9\" customHeight=\"1\"/><row r=\"4\" ht=\"42\" customHeight=\"1\" x14ac:dyDescent=\"0.25\">");
        });

        var edited = EditBytes(patched, p =>
        {
            var s = p.Workbook.Worksheets.Edit("S");
            Assert.Equal("Keep bold", s.Cells["A2"].Value);
            Assert.Equal((4, 1), s.Cells.Dimension);
            s.Cells["B4"].Value = "added";
        });

        using var doc = OpenValidated(edited);
        var ws = doc.WorkbookPart.WorksheetParts.Single().Worksheet;
        var rows = ws.Descendants<X.Row>().ToDictionary(r => r.RowIndex.Value);
        Assert.True(rows[1].Hidden.Value);
        Assert.Equal((byte)1, rows[2].OutlineLevel.Value);
        Assert.Equal(9, rows[3].Height.Value);                                   // Empty row kept for its height.
        Assert.Equal(42, rows[4].Height.Value);
        Assert.Equal("0.25", rows[4].DyDescent.InnerText);
        var cells = ws.Descendants<X.Cell>().ToDictionary(c => c.CellReference.Value);
        Assert.Equal(2, cells["A2"].InlineString.Elements<X.Run>().Count());    // Rich runs intact.
        Assert.NotNull(cells["A2"].InlineString.Elements<X.Run>().Last().RunProperties.GetFirstChild<X.Bold>());
        Assert.Equal("added", doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ElementAt(int.Parse(cells["B4"].CellValue.Text)).InnerText);
    }
}
