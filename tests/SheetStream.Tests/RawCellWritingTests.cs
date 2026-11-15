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

    /// <summary>The staging buffer's size, as the boundary scripts need it. Kept in step with the writer by the test below.</summary>
    private const int BufferLength = 1 << 14;

    private sealed class XNamespaceHolder
    {
        public readonly string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public readonly string Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    }

    [Fact]
    public void TheTestsKnowTheBufferSize() => Assert.Equal(BufferLength, RowWriter.RawCapacity);

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
    /// Random cells and random nasty text from fixed seeds (never <c>string.GetHashCode</c>, which differs from one
    /// process to the next), written through both paths with every combination of options.
    /// </summary>
    [Fact]
    public void RandomCellsAndText_AreWrittenExactlyAsXmlWriterWroteThem()
    {
        for (var seed = 1; seed <= 60; seed++)
            AssertSameAsReference(RawCellScript.Fuzz(20261115 + seed, 150), shared: seed % 3 != 0, references: seed % 2 == 0, autoFit: seed % 5 == 0);
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

    private static void AssertSameAsReference(RawCellScript script, bool shared, bool references, bool autoFit)
    {
        string expected = null!;
        var bytes = ExcelPackage.CreateBytes(wb =>
        {
            var styles = RawCellScript.CreateStyles(wb);
            var sheet = wb.Worksheets.Add("S");
            if (autoFit) sheet.AutoFitColumns();
            script.Apply(sheet, styles);
            expected = Reference(script, styles, wb.Styles, shared, references);
        }, new PackageOptions { UseSharedStrings = shared, WriteCellReferences = references, CompressionLevel = CompressionLevel.Fastest });

        var actual = SheetData(bytes);
        var at = FirstDifference(expected, actual);
        Assert.True(at < 0, $"Output differs from the XmlWriter reference at character {at} (shared={shared}, references={references}, autoFit={autoFit}).\nExpected: {Around(expected, at)}\n  Actual: {Around(actual, at)}");
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
    private static string Reference(RawCellScript script, CellStyle[] styles, StyleSheet styleSheet, bool shared, bool references)
    {
        var output = new MemoryStream();
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            CloseOutput = false,
            Indent = false,
            NewLineHandling = NewLineHandling.None,
            OmitXmlDeclaration = false,
        };
        var sharedIndex = new Dictionary<string, int>(StringComparer.Ordinal);
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
