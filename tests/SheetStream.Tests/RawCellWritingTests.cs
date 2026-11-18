using System.IO.Compression;
using System.Text;
using System.Xml;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream.Tests;

/// <summary>
/// Cells of the common kinds are composed as text in a buffer and handed to the XML writer in one piece, which
/// does no escaping and no checking of its own. These tests pin that the file is exactly what it was when every
/// cell went through <c>XmlWriter.WriteStartElement</c>, <c>WriteAttributeString</c> and <c>WriteString</c>: a
/// reference implementation of the old writing code, built on a real <see cref="XmlWriter"/> with the library's
/// settings, writes the same script of cells and the two outputs must match character for character.
/// </summary>
public class RawCellWritingTests
{
    private static readonly XNamespaceHolder Ns = new();

    /// <summary>
    /// The size of the buffer rows are composed in (the write buffer of the zip entry, <see cref="PackageOptions.BufferSize"/>),
    /// as the boundary scripts need it. Kept in step with the option's default by the test below.
    /// </summary>
    private const int BufferLength = 64 * 1024;

    private sealed class XNamespaceHolder
    {
        public readonly string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public readonly string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    }

    [Fact]
    public void TheTestsKnowTheBufferSize() => Assert.Equal(BufferLength, new PackageOptions().BufferSize);

    // ---- The differential tests ---------------------------------------------------------------------------------

    public static IEnumerable<object[]> Options()
    {
        foreach (var shared in new[] { true, false })
            foreach (var references in new[] { false, true })
                foreach (var autoFit in new[] { false, true })
                    yield return new object[] { shared, references, autoFit };
    }

    [Theory]
    [MemberData(nameof(Options))]
    public void CuratedValues_AreWrittenExactlyAsXmlWriterWroteThem(bool shared, bool references, bool autoFit)
        => AssertSameAsReference(RawCellScript.Curated(), shared, references, autoFit);

    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, true, true)]
    public void EveryCharacter_IsWrittenExactlyAsXmlWriterWroteIt(bool shared, bool references, bool autoFit)
        => AssertSameAsReference(RawCellScript.CharacterSweep(), shared, references, autoFit);

    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, true, true)]
    public void TextAroundTheBufferSize_IsNeverCutOrReordered(bool shared, bool references, bool autoFit)
        => AssertSameAsReference(RawCellScript.BufferBoundaries(BufferLength), shared, references, autoFit);

    /// <summary>
    /// The smallest write buffer there is (4 KB, whatever <see cref="PackageOptions.BufferSize"/> asks for), so that every
    /// flush lands somewhere new: inside text, inside an entity, between a cell's pieces and between rows.
    /// </summary>
    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, true)]
    public void ASmallBuffer_NeverCutsOrReordersAnything(bool shared, bool references, bool autoFit)
    {
        AssertSameAsReference(RawCellScript.BufferBoundaries(4096), shared, references, autoFit, bufferSize: 100);
        AssertSameAsReference(RawCellScript.Curated(), shared, references, autoFit, bufferSize: 100);
        AssertSameAsReference(RawCellScript.Fuzz(7, 400), shared, references, autoFit, bufferSize: 100);
    }

    /// <summary>
    /// A sheet filled into a template is written by the same cell code but around an XML writer that also copies the
    /// template's elements, so its rows are handed over as text (and the row element is the XML writer's). Its cells
    /// must be the same too.
    /// </summary>
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void ASheetFilledIntoATemplate_IsWrittenExactlyAsXmlWriterWroteIt(bool shared, bool references)
    {
        AssertSameAsReferenceThroughTemplate(RawCellScript.Curated(), shared, references);
        AssertSameAsReferenceThroughTemplate(RawCellScript.BufferBoundaries(RowWriter.XmlTargetBufferLength), shared, references);
        for (var seed = 1; seed <= 6; seed++)
            AssertSameAsReferenceThroughTemplate(RawCellScript.Fuzz(900 + seed, 120), shared, references);
    }

    /// <summary>
    /// Random cells and random nasty text from fixed seeds (never <c>string.GetHashCode</c>, which differs from one
    /// process to the next), written through both paths with every combination of options.
    /// </summary>
    [Fact]
    public void RandomCellsAndText_AreWrittenExactlyAsXmlWriterWroteThem()
    {
        for (var seed = 1; seed <= 60; seed++)
            AssertSameAsReference(RawCellScript.Fuzz(20261115 + seed, 150), shared: seed % 3 != 0, references: seed % 2 == 0, autoFit: seed % 5 == 0);
    }

    /// <summary>
    /// Inline text is written in one pass, eight characters at a time where it can be, with a separate route for
    /// anything that needs sanitising. Every length around the block size and the size limit of the one-pass route,
    /// with every kind of special character at every position (and at none), must still be exactly what the XML
    /// writer wrote, with and without a style and an address on the cell.
    /// </summary>
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void InlineTextOfEveryLength_WithASpecialCharacterAtEveryPosition_IsWrittenAsXmlWriterWroteIt(bool references, bool styled)
    {
        var specials = new[] { "&", "<", ">", "_", "_x0041_", "é", "€", "😀", "\u0001", "\u001F", " ", "\t", "\r\n", "￾", "\ud800", "\udc00", "\"", "\u007F", "ñ&" };
        var script = new RawCellScript();
        foreach (var length in new[] { 0, 1, 2, 7, 8, 9, 15, 16, 17, 23, 24, 25, 40, 63, 64, 65 })
        {
            script.NewRow().Add(new ScriptCell { Kind = ScriptKind.Text, Text = new string('a', length), Style = styled ? 2 : -1 });
            foreach (var special in specials)
                for (var at = 0; at <= length; at++)
                {
                    var text = new string('a', at) + special + new string('b', length - at);
                    script.NewRow().Add(new ScriptCell { Kind = ScriptKind.Text, Text = text, Style = styled ? 2 : -1 });
                }
        }
        AssertSameAsReference(script, shared: false, references, autoFit: false);
    }

    /// <summary>The one-pass route takes text up to 512 characters; longer text, and text in a buffer too small for it, must give the same bytes.</summary>
    [Theory]
    [InlineData(65536)]
    [InlineData(100)]
    public void InlineTextAroundTheSizeLimitOfTheOnePassRoute_IsWrittenAsXmlWriterWroteIt(int bufferSize)
    {
        var script = new RawCellScript();
        foreach (var length in new[] { 510, 511, 512, 513, 514, 700, 1200, 5000 })
            foreach (var tail in new[] { "", "&", "<>", "é", "😀", "_x0041_", "\u0001" })
            {
                var text = new string('x', length) + tail;
                var row = script.NewRow();
                row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = text, Style = -1 });
                row.Add(new ScriptCell { Kind = ScriptKind.Text, Text = " " + text + " ", Style = 3 });
                row.Add(new ScriptCell { Kind = ScriptKind.Int, Integer = length, Style = -1 });
            }
        AssertSameAsReference(script, shared: false, references: false, autoFit: false, bufferSize: bufferSize);
    }

    /// <summary>
    /// <see cref="StringStorage.Auto"/>: columns that are shared, inline, sampled, shared and then let go, and cells that
    /// must be kept out of inline text, mixed with numbers, skipped columns and styles, over enough rows for every
    /// column to be judged, must be written exactly as the reference writes them.
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AutoMode_MixedAndChangingColumns_AreWrittenExactlyAsXmlWriterWroteThem(bool references)
    {
        var phrases = new[] { "Pending approval by the finance team", "Waiting for the warehouse to confirm stock", "Customer asked for a call back" };
        var random = new Random(4711);
        var script = new RawCellScript();
        for (var r = 0; r < 9000; r++)
        {
            var row = script.NewRow();
            // 1: shared, then every text is new (let go at the end of a monitoring window).
            row.Add(Text(r < 1200 ? phrases[r % 3] : "free text " + r, -1));
            // 2: unique from the start (inline after the early exit).
            row.Add(Text("ref " + r + (r % 50 == 0 ? " & <more>" : ""), r % 9 == 0 ? 3 : -1));
            row.Add(new ScriptCell { Kind = ScriptKind.Int, Integer = r, Style = -1 });
            // 3: many distinct short values (shared on the count of distinct values).
            row.Add(Text("N" + random.Next(1500), -1));
            // 4: a short, small set (inline after the first sample).
            row.Add(Text(new[] { "New", "Pending", "Shipped" }[r % 3], -1));
            if (r % 7 == 0) row.Add(new ScriptCell { Kind = ScriptKind.Skip, Integer = 2, Style = -1 });
            // 5: text that Sanitize changes, in a column that is otherwise inline, and text that is plain.
            row.Add(Text(r % 11 == 0 ? "control \u0001 and _x0041_ " + r : "plain " + r + " é €", -1));
            row.Add(new ScriptCell { Kind = ScriptKind.Empty, Style = r % 13 == 0 ? 4 : -1 });
            row.Add(Text(r % 5 == 0 ? null : " padded " + r + " ", -1));
        }
        AssertSameAsReference(script, shared: false, references, autoFit: false, auto: true);

        static ScriptCell Text(string? text, int style) => new ScriptCell { Kind = text == null ? ScriptKind.NullText : ScriptKind.Text, Text = text, Style = style };
    }
    /// <summary>The row writer's own rows and the buffered cells meet at the row end: a long row, then rows of different shapes.</summary>
    [Fact]
    public void ManyRowsOfOneKind_FillTheBufferAcrossRows()
    {
        var script = new RawCellScript();
        for (var r = 0; r < 4000; r++)
        {
            var row = script.NewRow();
            for (var c = 0; c < 10; c++)
                row.Add(c % 2 == 0
                    ? new ScriptCell { Kind = ScriptKind.Double, Number = r * 1.25 + c, Style = c == 4 ? 3 : -1 }
                    : new ScriptCell { Kind = ScriptKind.Text, Text = "item " + (r + c) % 50, Style = -1 });
        }
        AssertSameAsReference(script, shared: true, references: false, autoFit: false);
    }

    // ---- Edge behaviour of the batching itself ------------------------------------------------------------------

    /// <summary>
    /// A cell written through the XmlWriter calls (a formula, rich text, an error) must land between the cells
    /// composed in the buffer, not before or after them.
    /// </summary>
    [Fact]
    public void CellsOnTheSlowPath_KeepTheirPlaceBetweenBufferedCells()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteCell(1).WriteFormula("A1*2").WriteCell("text").WriteRichText(new RichText("rich") { Bold = true })
                .WriteCell(double.NaN).WriteCell(true).WriteFormula("A1+1", 3.0).WriteCell("tail & <more>");
            s.Row().WriteRichText(new RichText("only rich")).WriteFormula("1+1");
            s.Row().WriteCell(2);
        });

        var data = SheetData(bytes);
        var order = new[]
        {
            "<c><v>1</v></c>", "<f>A1*2</f>", "<c t=\"s\"><v>0</v></c>", "<is><r>", "t=\"e\"><v>#NUM!</v>", "t=\"b\"><v>1</v>", "<f>A1+1</f><v>3</v>",
            "<c t=\"s\"><v>1</v></c>", "only rich", "<f>1+1</f>", "<c><v>2</v></c>",
        };
        var at = 0;
        foreach (var piece in order)
        {
            var found = data.IndexOf(piece, at, StringComparison.Ordinal);
            Assert.True(found >= 0, $"'{piece}' is missing or out of order in: {data}");
            at = found;
        }
    }

    /// <summary>
    /// The row element is written by hand for new sheets and by the XML writer for template sheets; both must give the
    /// attributes in the order XmlWriter always wrote them, and an empty row must close itself with " />".
    /// </summary>
    [Fact]
    public void RowOptions_AreWrittenInTheSameOrderAsBefore()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var tint = wb.Styles.Create(o => o.Fill.BackgroundColor = "#FFEEDD");
            var s = wb.Worksheets.Add("S");
            s.Row(1, new RowOptions { Style = tint, Height = 20.5, Hidden = true, OutlineLevel = 2, Collapsed = true }).WriteCell(1);
            s.Row(2, null, 15);                                   // An empty row with a height.
            s.Row(3, new RowOptions { OutlineLevel = 1 }).WriteCell("x");
            s.Row(4, tint).WriteCell(2.5);                        // The row style reaches the cell.
        });

        var data = SheetData(bytes);
        Assert.StartsWith("<sheetData><row r=\"1\" s=\"1\" customFormat=\"1\" ht=\"20.5\" customHeight=\"1\" hidden=\"1\" outlineLevel=\"2\" collapsed=\"1\"><c s=\"1\"><v>1</v></c></row>", data);
        Assert.Contains("<row r=\"2\" ht=\"15\" customHeight=\"1\" />", data);
        Assert.Contains("<row r=\"3\" outlineLevel=\"1\"><c t=\"s\"><v>0</v></c></row>", data);
        Assert.Contains("<row r=\"4\" s=\"1\" customFormat=\"1\"><c s=\"1\"><v>2.5</v></c></row>", data);
    }

    [Fact]
    public void ARowHeightOutOfRange_IsRefusedWithoutBreakingTheSheet()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row(1).WriteCell(1);
            Assert.Throws<ArgumentOutOfRangeException>(() => s.Row(2, null, 500));
            s.Row(3).WriteCell(3);
        });
        Assert.DoesNotContain("<row r=\"2\"", SheetData(bytes));
    }

    [Fact]
    public void AnEmptyRow_StillClosesItself()
    {
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var s = wb.Worksheets.Add("S");
            s.Row().WriteCell(1);
            s.Row();
            s.Row().Skip(3);
            s.Row().WriteEmpty();
        });
        Assert.Equal("<sheetData><row r=\"1\"><c><v>1</v></c></row><row r=\"2\" /><row r=\"3\" /><row r=\"4\" /></sheetData>", SheetData(bytes));
    }

    // ---- The reference: the cell-writing code as it was before the buffer, on a real XmlWriter --------------------

    private static void AssertSameAsReference(RawCellScript script, bool shared, bool references, bool autoFit, int bufferSize = BufferLength, bool auto = false)
    {
        string expected = null!;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var styles = RawCellScript.CreateStyles(wb);
            var sheet = wb.Worksheets.Add("S");
            if (autoFit) sheet.AutoFitColumns();
            script.Apply(sheet, styles);
            expected = Reference(script, styles, wb.Styles, shared, references, auto);
        }, new PackageOptions
        {
            StringStorage = auto ? StringStorage.Auto : shared ? StringStorage.Shared : StringStorage.Inline,
            WriteCellReferences = references,
            CompressionLevel = CompressionLevel.Fastest,
            BufferSize = bufferSize,
            // The reference writes every distinct string to the table, so the memory budget must not turn any inline.
            MaxSharedStringChars = long.MaxValue / 2,
        });

        var actual = SheetData(bytes);
        var at = FirstDifference(expected, actual);
        Assert.True(at < 0, $"Output differs from the XmlWriter reference at character {at} (shared={shared}, references={references}, autoFit={autoFit}).\nExpected: {Around(expected, at)}\n  Actual: {Around(actual, at)}");
    }

    private static void AssertSameAsReferenceThroughTemplate(RawCellScript script, bool shared, bool references)
    {
        var template = ExcelPackage.CreateBytes(wb => wb.Worksheets.Add("S"));
        var options = new PackageOptions
        {
            StringStorage = shared ? StringStorage.Shared : StringStorage.Inline,
            WriteCellReferences = references,
            CompressionLevel = CompressionLevel.Fastest,
            MaxSharedStringChars = long.MaxValue / 2,
        };
        string expected = null!;
        using var output = new MemoryStream();
        using (var package = ExcelPackage.OpenTemplate(new MemoryStream(template), output, leaveOpen: true, options))
        {
            var styles = RawCellScript.CreateStyles(package.Workbook);
            script.Apply(package.Workbook.Worksheets.Open("S"), styles);
            expected = Reference(script, styles, package.Workbook.Styles, shared, references);
            package.Save();
        }

        var actual = SheetData(output.ToArray());
        var at = FirstDifference(expected, actual);
        Assert.True(at < 0, $"Output differs from the XmlWriter reference at character {at} (template, shared={shared}, references={references}).\nExpected: {Around(expected, at)}\n  Actual: {Around(actual, at)}");
    }

    private static int FirstDifference(string a, string b)
    {
        var n = Math.Min(a.Length, b.Length);
        for (var i = 0; i < n; i++) if (a[i] != b[i]) return i;
        return a.Length == b.Length ? -1 : n;
    }

    private static string Around(string text, int at)
    {
        var from = Math.Max(0, at - 120);
        return text.Substring(from, Math.Min(240, text.Length - from)).Replace("\n", "\\n").Replace("\r", "\\r");
    }

    /// <summary>The text of the sheet's <c>sheetData</c> element as the package wrote it.</summary>
    private static string SheetData(byte[] workbook)
    {
        using var zip = new ZipArchive(new MemoryStream(workbook), ZipArchiveMode.Read);
        using var reader = new StreamReader(zip.GetEntry("xl/worksheets/sheet1.xml")!.Open(), new UTF8Encoding(false));
        return Extract(reader.ReadToEnd());
    }

    private static string Extract(string part)
    {
        var start = part.IndexOf("<sheetData", StringComparison.Ordinal);
        var end = part.IndexOf("</sheetData>", StringComparison.Ordinal);
        Assert.True(start >= 0 && end > start, "sheetData not found");
        return part.Substring(start, end + "</sheetData>".Length - start);
    }

    /// <summary>
    /// Writes the script the way <c>RowWriter</c> did before cells were composed in a buffer: one XmlWriter call per
    /// element, attribute and piece of text, with the library's writer settings and the same wrapping elements.
    /// Everything else (when the address is written, the style rules, the sanitising of text) is the library's own
    /// code, so what this pins is the XML writing, which is all that changed.
    /// </summary>
    private static string Reference(RawCellScript script, CellStyle[] styles, StyleSheet styleSheet, bool shared, bool references, bool auto = false)
    {
        var output = new MemoryStream();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            CloseOutput = false,
            Indent = false,
            // Entitize writes a carriage return in text as &#xD; (a raw one would be read back as LF) and leaves tabs and
            // line feeds in text alone, which is what the library writes. The cells' attributes carry no such characters.
            NewLineHandling = NewLineHandling.Entitize,
            OmitXmlDeclaration = false,
        };
        var sharedIndex = new Dictionary<string, int>(StringComparer.Ordinal);
        // StringStorage.Auto: the reference keeps the same per-column judgement (StringColumnState has its own unit tests) and writes the cells with plain XmlWriter calls.
        var columnStates = new StringColumnState[CellReference.MaxColumns + 1];
        using (var w = XmlWriter.Create(output, settings))
        {
            w.WriteStartDocument(true);
            w.WriteStartElement("worksheet", Ns.Main);
            w.WriteAttributeString("xmlns", "r", null, Ns.Relationships);
            w.WriteStartElement("sheetData");

            var rowIndex = 0;
            foreach (var row in script.Rows)
            {
                rowIndex++;
                w.WriteStartElement("row");
                w.WriteAttributeString("r", rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
                var next = 1;
                var last = 0;

                void Begin(CellStyle? style, string? type)
                {
                    w.WriteStartElement("c");
                    if (references || next != last + 1)
                        w.WriteAttributeString("r", CellReference.ColumnLetters(next) + rowIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    last = next;
                    if (style != null && !style.IsDefault) w.WriteAttributeString("s", style.Id.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    if (type != null) w.WriteAttributeString("t", type);
                }

                void Value(Action<XmlWriter> write)
                {
                    w.WriteStartElement("v");
                    write(w);
                    w.WriteEndElement();
                    w.WriteEndElement();
                    next++;
                }

                void Empty(CellStyle? style)
                {
                    if (style != null && !style.IsDefault)
                    {
                        Begin(style, null);
                        w.WriteEndElement();
                    }
                    next++;
                }

                void Inline(string value, CellStyle? style)
                {
                    Begin(style, "inlineStr");
                    w.WriteStartElement("is");
                    w.WriteStartElement("t");
                    var text = XmlTextUtil.Sanitize(value);
                    if (XmlTextUtil.NeedsSpacePreserve(text)) w.WriteAttributeString("xml", "space", null, "preserve");
                    w.WriteString(text);
                    w.WriteEndElement();
                    w.WriteEndElement();
                    w.WriteEndElement();
                    next++;
                }

                // The formula text as the library prepares it: trimmed, without a leading '=', sanitised.
                void Formula(string text)
                {
                    var body = text.Trim();
                    if (body[0] == '=') body = body.Substring(1);
                    w.WriteStartElement("f");
                    w.WriteString(XmlTextUtil.Sanitize(body));
                    w.WriteEndElement();
                }

                void Number(double value, CellStyle? style)
                {
                    if (double.IsNaN(value) || double.IsInfinity(value))
                    {
                        Begin(style, "e");
                        Value(x => x.WriteString("#NUM!"));
                        return;
                    }
                    Begin(style, null);
                    Value(x => x.WriteValue(value));
                }

                foreach (var cell in row)
                {
                    var style = cell.Style < 0 ? null : styles[cell.Style];
                    switch (cell.Kind)
                    {
                        case ScriptKind.Int: Begin(style, null); Value(x => x.WriteValue((int)cell.Integer)); break;
                        case ScriptKind.Long: Begin(style, null); Value(x => x.WriteValue(cell.Integer)); break;
                        case ScriptKind.Double: Number(cell.Number, style); break;
                        case ScriptKind.Float: Number((double)(float)cell.Number, style); break;
                        case ScriptKind.Decimal: Begin(style, null); Value(x => x.WriteValue(cell.Money)); break;
                        case ScriptKind.Date:
                        {
                            var effective = style ?? (cell.Date.TimeOfDay == TimeSpan.Zero ? styleSheet.DateStyle : styleSheet.DateTimeStyle);
                            Begin(effective, null);
                            Value(x => x.WriteValue(ExcelDate.ToSerial(cell.Date)));
                            break;
                        }
                        case ScriptKind.TimeSpan:
                            Begin(style ?? styleSheet.DurationStyle, null);
                            Value(x => x.WriteValue(cell.Span.TotalDays));
                            break;
                        case ScriptKind.Bool: Begin(style, "b"); Value(x => x.WriteString(cell.Integer != 0 ? "1" : "0")); break;
                        case ScriptKind.Text when auto:
                        {
                            ref var state = ref columnStates[next];
                            var inlineColumn = state.WritesInline(); // Also counts a cell of a column that is waiting for its next look.
                            var sanitized = XmlTextUtil.Sanitize(cell.Text!);
                            if (inlineColumn && ReferenceEquals(sanitized, cell.Text)) { Inline(cell.Text!, style); break; }
                            // Sampled, looking or shared columns, and text that had to be changed (kept out of inline cells, see RowWriter.WriteInlineCore).
                            var known = sharedIndex.TryGetValue(sanitized, out var index);
                            if (!known) sharedIndex[sanitized] = index = sharedIndex.Count;
                            if (!inlineColumn) state.Observe(known, cell.Text!.Length);
                            Begin(style, "s");
                            Value(x => x.WriteValue(index));
                            break;
                        }
                        case ScriptKind.Text when shared:
                        {
                            var sanitized = XmlTextUtil.Sanitize(cell.Text!);
                            if (!sharedIndex.TryGetValue(sanitized, out var index)) sharedIndex[sanitized] = index = sharedIndex.Count;
                            Begin(style, "s");
                            Value(x => x.WriteValue(index));
                            break;
                        }
                        case ScriptKind.Text: Inline(cell.Text!, style); break;
                        case ScriptKind.Inline: Inline(cell.Text!, style); break;
                        case ScriptKind.NullText: Empty(style); break;
                        case ScriptKind.Empty: Empty(style); break;
                        case ScriptKind.Skip: next += (int)cell.Integer; break;
                        case ScriptKind.Formula: Begin(style, null); Formula(cell.Text!); w.WriteEndElement(); next++; break;
                        case ScriptKind.FormulaNumber:
                            Begin(style, null);
                            Formula(cell.Text!);
                            w.WriteStartElement("v");
                            w.WriteValue(cell.Number);
                            w.WriteEndElement();
                            w.WriteEndElement();
                            next++;
                            break;
                        case ScriptKind.FormulaText:
                            Begin(style, "str");
                            Formula(cell.Text!);
                            w.WriteElementString("v", XmlTextUtil.Sanitize(cell.Text2!));
                            w.WriteEndElement();
                            next++;
                            break;
                        case ScriptKind.FormulaBool:
                            Begin(style, "b");
                            Formula(cell.Text!);
                            w.WriteElementString("v", cell.Integer != 0 ? "1" : "0");
                            w.WriteEndElement();
                            next++;
                            break;
                        case ScriptKind.RichText:
                            Begin(style, "inlineStr");
                            w.WriteStartElement("is");
                            foreach (var run in RawCellScript.Runs(cell)) run.WriteTo(w);
                            w.WriteEndElement();
                            w.WriteEndElement();
                            next++;
                            break;
                        default: throw new InvalidOperationException(cell.Kind.ToString());
                    }
                }
                w.WriteEndElement();
            }

            w.WriteEndElement(); // sheetData
            w.WriteEndElement(); // worksheet
            w.WriteEndDocument();
        }
        return Extract(new UTF8Encoding(false).GetString(output.ToArray()));
    }
}
