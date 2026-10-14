using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using SheetStream.Internal;
using SheetStream.Reading;
using SheetStream.Styles;

namespace SheetStream
{
    public sealed partial class Worksheet
    {
        // Edit mode: a template sheet whose rows were all loaded into the random-access buffer. Saving writes
        // the buffer in place of the original sheetData; everything around it is copied as in template mode.
        private bool _editMode;
        private int _editedLastRow;                                   // Last row number present in the original sheet.
        private SortedDictionary<int, List<RawAttribute>>? _rowAttributes; // Original <row> attributes, by row.

        /// <summary>
        /// Counts the cells of a template sheet before it is opened for editing, so a sheet over
        /// <see cref="PackageOptions.MaxEditableCells"/> is refused before the package changes state.
        /// </summary>
        internal static void EnsureEditable(TemplateSource template, string part, string sheetName, int limit)
        {
            var count = 0;
            using (var archive = template.OpenArchive())
            using (var r = template.OpenXml(archive, part) ?? throw new InvalidOperationException($"The workbook has no part '{part}' for sheet '{sheetName}'."))
            {
                while (r.Read())
                {
                    if (r.NodeType != XmlNodeType.Element || r.LocalName != "c") continue;
                    if (++count > limit) ThrowTooLarge(sheetName, limit);
                    if (!r.IsEmptyElement) r.Skip(); // Nothing below a cell is a cell.
                }
            }
        }

        private static void ThrowTooLarge(string sheetName, int limit)
            => throw new InvalidOperationException($"Sheet '{sheetName}' has more than {limit:N0} cells. Editing keeps the whole sheet in memory and is meant for small workbooks; for a sheet this size stream the rows into a copy with Worksheets.Open (template mode) or write it from scratch with Row(). The limit is PackageOptions.MaxEditableCells.");

        /// <summary>
        /// Loads every row and cell of the template sheet into the random-access buffer. Each cell keeps its
        /// original XML for a faithful round trip and gets a parsed value, style and formula for reading.
        /// Must run right after <see cref="AttachTemplate"/>, before anything is written.
        /// </summary>
        internal void LoadTemplateCells()
        {
            if (_templateSource == null) throw new InvalidOperationException("Only a template sheet can be loaded for editing.");
            _editMode = true;
            _buffer = new SortedDictionary<int, SortedDictionary<int, BufferedCell>>();
            _cells ??= new CellCollection(this);

            var limit = Workbook.Package.Options.MaxEditableCells;
            var context = new LoadContext(_templateSource.GetSharedStringTexts(), Workbook.Styles.CustomNumberFormatCodes(), _templateSource.Date1904);
            var count = 0;
            var lastRow = 0;
            using (var archive = _templateSource.OpenArchive())
            using (var r = _templateSource.OpenXml(archive, TemplatePart!)!)
            {
                while (r.Read())
                {
                    if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "sheetData") break;
                    if (r.NodeType != XmlNodeType.Element || r.LocalName != "row") continue;
                    lastRow = LoadRow(r, lastRow, context, limit, ref count);
                }
            }
        }

        /// <summary>Everything the cell loader needs to turn stored text into typed values.</summary>
        private sealed class LoadContext
        {
            public LoadContext(IReadOnlyList<string> sharedStrings, Dictionary<int, string> numberFormatCodes, bool date1904)
            {
                SharedStrings = sharedStrings;
                NumberFormatCodes = numberFormatCodes;
                Date1904 = date1904;
            }

            public IReadOnlyList<string> SharedStrings { get; }
            public Dictionary<int, string> NumberFormatCodes { get; }
            public bool Date1904 { get; }
            public StringBuilder Inner { get; } = new StringBuilder();
        }

        /// <summary>Loads one <c>row</c> element (reader on its start tag) and returns its row number.</summary>
        private int LoadRow(XmlReader r, int lastRow, LoadContext context, int limit, ref int count)
        {
            var rowIndex = TemplateSource.Int(r.GetAttribute("r")) ?? lastRow + 1;
            if (rowIndex < 1 || rowIndex > CellReference.MaxRows)
                throw new System.IO.InvalidDataException($"Sheet '{Name}' has a row numbered {rowIndex}, outside Excel's range. The file is damaged.");
            if (rowIndex > _editedLastRow) _editedLastRow = rowIndex;

            // Keep the row's own attributes (height, hidden, outline level, x14ac:dyDescent...) verbatim. The
            // number is rewritten by the row writer and spans is only a hint that the new cells may invalidate.
            List<RawAttribute>? attributes = null;
            if (r.MoveToFirstAttribute())
            {
                do
                {
                    if (r.Prefix == "xmlns" || (r.Prefix.Length == 0 && (r.LocalName == "xmlns" || r.LocalName == "r" || r.LocalName == "spans"))) continue;
                    (attributes ??= new List<RawAttribute>()).Add(new RawAttribute(r.Prefix, r.LocalName, r.NamespaceURI, r.Value));
                }
                while (r.MoveToNextAttribute());
                r.MoveToElement();
            }
            if (attributes != null)
            {
                (_rowAttributes ??= new SortedDictionary<int, List<RawAttribute>>())[rowIndex] = attributes;
                if (!_buffer!.ContainsKey(rowIndex)) _buffer[rowIndex] = new SortedDictionary<int, BufferedCell>();
            }
            if (r.IsEmptyElement) return rowIndex;

            var depth = r.Depth;
            var nextColumn = 1;
            r.Read();
            while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == depth))
            {
                if (r.NodeType == XmlNodeType.Element && r.LocalName == "c")
                {
                    if (++count > limit) ThrowTooLarge(Name, limit);
                    LoadCell(r, rowIndex, ref nextColumn, context); // Leaves the reader after the cell.
                    continue;
                }
                if (r.NodeType == XmlNodeType.Element) r.Skip(); else r.Read();
            }
            return rowIndex;
        }

        /// <summary>
        /// Loads one <c>c</c> element (reader on its start tag) into the buffer and leaves the reader on the
        /// node after it. The walk serialises the cell's inner XML and picks out the formula, cached value and
        /// inline text in the same pass, so a cell is parsed once.
        /// </summary>
        private void LoadCell(XmlReader r, int rowIndex, ref int nextColumn, LoadContext context)
        {
            var reference = r.GetAttribute("r");
            var column = reference != null ? ExcelReader.ColumnOf(reference) : nextColumn;
            nextColumn = column + 1;

            string? type = null;
            var styleIndex = 0;
            List<RawAttribute>? extra = null;
            if (r.MoveToFirstAttribute())
            {
                do
                {
                    if (r.Prefix == "xmlns" || (r.Prefix.Length == 0 && r.LocalName == "xmlns")) continue;
                    if (r.Prefix.Length == 0 && r.LocalName == "r") continue;
                    if (r.Prefix.Length == 0 && r.LocalName == "s") { styleIndex = TemplateSource.Int(r.Value) ?? 0; continue; }
                    if (r.Prefix.Length == 0 && r.LocalName == "t") { type = r.Value; continue; }
                    (extra ??= new List<RawAttribute>()).Add(new RawAttribute(r.Prefix, r.LocalName, r.NamespaceURI, r.Value));
                }
                while (r.MoveToNextAttribute());
                r.MoveToElement();
            }
            if (styleIndex < 0 || !Workbook.Styles.IsValidStyleId(styleIndex))
                throw new System.IO.InvalidDataException($"Cell {CellReference.ToA1(rowIndex, column)} on sheet '{Name}' uses style index {styleIndex}, which the workbook's styles part does not define. The file is damaged.");

            var cell = new BufferedCell
            {
                RawType = type,
                RawAttributes = extra,
                Style = styleIndex > 0 ? new CellStyle(styleIndex) : null,
            };

            string? v = null;
            string? formula = null;
            string? inline = null;
            if (r.IsEmptyElement)
            {
                cell.RawInner = "";
                r.Read();
            }
            else
            {
                WalkCellContent(r, cell, context, out v, out formula, out inline);
            }

            cell.Formula = formula;
            cell.Value = TypedValue(type, v, inline, styleIndex, context);

            // Duplicate references in a damaged file: the later cell wins, as it does in Excel's repair.
            if (!_buffer!.TryGetValue(rowIndex, out var cols)) _buffer[rowIndex] = cols = new SortedDictionary<int, BufferedCell>();
            cols[column] = cell;
        }

        /// <summary>
        /// Walks the children of a non-empty <c>c</c> element, copying them to the cell's raw XML and capturing
        /// <c>f</c>, <c>v</c> and inline <c>t</c> text. The reader ends on the node after the cell's end tag.
        /// </summary>
        private static void WalkCellContent(XmlReader r, BufferedCell cell, LoadContext context, out string? v, out string? formula, out string? inline)
        {
            var sb = context.Inner;
            sb.Clear();
            v = null;
            formula = null;
            inline = null;
            var vText = (StringBuilder?)null;
            var fText = (StringBuilder?)null;
            var inlineText = (StringBuilder?)null;

            var cellDepth = r.Depth;
            var childDepth = cellDepth + 1;
            var capture = (StringBuilder?)null; // Where text nodes go while inside f, v or an inline t.
            var captureDepth = -1;
            var inInline = false;
            var phoneticDepth = -1;              // Inside rPh: copied, but not part of the cell text.
            r.Read();
            while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == cellDepth))
            {
                switch (r.NodeType)
                {
                    case XmlNodeType.Element:
                    {
                        var local = r.LocalName;
                        var depth = r.Depth;
                        var empty = r.IsEmptyElement;
                        if (depth - cellDepth > XmlCopy.MaxDepth)
                            throw new System.IO.InvalidDataException($"A cell nests XML deeper than {XmlCopy.MaxDepth} levels, which no Excel file does. It is damaged or crafted.");
                        if (depth == childDepth && r.NamespaceURI == OpenXmlNamespaces.Main)
                        {
                            if (local == "f")
                            {
                                fText = new StringBuilder();
                                var kind = r.GetAttribute("t");
                                if (kind == "shared" && r.GetAttribute("ref") != null) cell.SharedFormulaMaster = true;
                                // A shared dependent (<f t="shared" si="n"/>) has no text of its own; report no formula.
                                if (kind == "shared" && r.GetAttribute("ref") == null) fText = null;
                                if (!empty && fText != null) { capture = fText; captureDepth = depth; }
                            }
                            else if (local == "v")
                            {
                                vText = new StringBuilder();
                                if (!empty) { capture = vText; captureDepth = depth; }
                            }
                            else if (local == "is")
                            {
                                inlineText = new StringBuilder();
                                inInline = !empty;
                            }
                        }
                        else if (inInline && r.NamespaceURI == OpenXmlNamespaces.Main)
                        {
                            if (local == "rPh" && phoneticDepth < 0) phoneticDepth = depth;
                            else if (local == "t" && phoneticDepth < 0 && !empty) { capture = inlineText; captureDepth = depth; }
                        }
                        AppendStartTag(sb, r, empty);
                        r.Read();
                        continue;
                    }
                    case XmlNodeType.EndElement:
                        if (r.Depth == captureDepth) { capture = null; captureDepth = -1; }
                        if (r.Depth == phoneticDepth) phoneticDepth = -1;
                        if (r.Depth == childDepth && r.LocalName == "is") inInline = false;
                        sb.Append("</").Append(QualifiedName(r)).Append('>');
                        r.Read();
                        continue;
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        capture?.Append(r.Value);
                        AppendEscapedText(sb, r.Value);
                        r.Read();
                        continue;
                    case XmlNodeType.Comment:
                        sb.Append("<!--").Append(r.Value).Append("-->");
                        r.Read();
                        continue;
                    default:
                        r.Read(); // Processing instructions and the like have no business inside a cell.
                        continue;
                }
            }
            r.Read(); // Past </c>.

            cell.RawInner = sb.ToString();
            v = vText?.ToString();
            formula = fText?.ToString();
            inline = inlineText?.ToString();
        }

        /// <summary>
        /// Appends the start tag of the element the reader is on. Elements in the main namespace are written
        /// unprefixed, because the sheet writer declares that namespace as the default; anything else carries
        /// its own declaration so the fragment stays self-contained.
        /// </summary>
        private static void AppendStartTag(StringBuilder sb, XmlReader r, bool empty)
        {
            sb.Append('<').Append(QualifiedName(r));
            if (r.NamespaceURI != OpenXmlNamespaces.Main && r.NamespaceURI.Length > 0)
            {
                sb.Append(r.Prefix.Length == 0 ? " xmlns=\"" : " xmlns:" + r.Prefix + "=\"");
                AppendEscapedAttribute(sb, r.NamespaceURI);
                sb.Append('"');
            }
            if (r.MoveToFirstAttribute())
            {
                do
                {
                    if (r.Prefix == "xmlns" || (r.Prefix.Length == 0 && r.LocalName == "xmlns")) continue; // Declared above as needed.
                    sb.Append(' ');
                    if (r.Prefix.Length > 0)
                    {
                        sb.Append(r.Prefix).Append(':');
                        // Any prefix other than the built-in xml: needs its namespace declared on this element.
                        if (r.Prefix != "xml")
                        {
                            sb.Append(r.LocalName).Append("=\"");
                            AppendEscapedAttribute(sb, r.Value);
                            sb.Append("\" xmlns:").Append(r.Prefix).Append("=\"");
                            AppendEscapedAttribute(sb, r.NamespaceURI);
                            sb.Append('"');
                            continue;
                        }
                    }
                    sb.Append(r.LocalName).Append("=\"");
                    AppendEscapedAttribute(sb, r.Value);
                    sb.Append('"');
                }
                while (r.MoveToNextAttribute());
                r.MoveToElement();
            }
            sb.Append(empty ? "/>" : ">");
        }

        /// <summary>Element name as written: bare for the main namespace, prefixed otherwise.</summary>
        private static string QualifiedName(XmlReader r)
            => r.NamespaceURI == OpenXmlNamespaces.Main || r.Prefix.Length == 0 ? r.LocalName : r.Prefix + ":" + r.LocalName;

        private static void AppendEscapedText(StringBuilder sb, string s)
        {
            foreach (var c in s)
            {
                switch (c)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    default: sb.Append(c); break;
                }
            }
        }

        private static void AppendEscapedAttribute(StringBuilder sb, string s)
        {
            foreach (var c in s)
            {
                switch (c)
                {
                    case '&': sb.Append("&amp;"); break;
                    case '<': sb.Append("&lt;"); break;
                    case '>': sb.Append("&gt;"); break;
                    case '"': sb.Append("&quot;"); break;
                    case '\t': sb.Append("&#9;"); break;
                    case '\n': sb.Append("&#10;"); break;
                    case '\r': sb.Append("&#13;"); break;
                    default: sb.Append(c); break;
                }
            }
        }

        /// <summary>
        /// Turns a cell's stored text into the value a caller expects, with the same rules as
        /// <see cref="ExcelReader"/>: shared and inline strings become text, numbers with a date format become
        /// <see cref="DateTime"/>, elapsed-time formats <see cref="TimeSpan"/>, errors <see cref="ExcelErrorValue"/>.
        /// </summary>
        private object? TypedValue(string? type, string? v, string? inline, int styleIndex, LoadContext context)
        {
            switch (type)
            {
                case "s":
                    var index = TemplateSource.Int(v);
                    return index.HasValue && index.Value >= 0 && index.Value < context.SharedStrings.Count ? context.SharedStrings[index.Value] : null;
                case "inlineStr":
                    return inline == null ? null : ExcelReader.Decode(inline);
                case "str":
                    return v == null ? null : ExcelReader.Decode(v);
                case "b":
                    return v == null ? (object?)null : v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
                case "e":
                    return v == null ? null : new ExcelErrorValue(v);
                case "d":
                    return v != null && DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var iso) ? iso : (object?)v;
                default:
                    if (v == null || !double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return v;
                    var kind = ExcelReader.Classify(Workbook.Styles.NumberFormatId(styleIndex), context.NumberFormatCodes);
                    if (kind == ExcelReader.CellKind.Date)
                    {
                        if (context.Date1904) number += 1462;
                        if (number >= -657435 && number < 2958466) return DateTime.FromOADate(number);
                    }
                    else if (kind == ExcelReader.CellKind.Duration)
                    {
                        return TimeSpan.FromDays(number);
                    }
                    return number;
            }
        }
    }
}
