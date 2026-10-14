using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream
{
    public sealed partial class Worksheet
    {
        // Template mode: the sheet streams the template's XML around the rows the caller writes.
        private TemplateSource? _templateSource;
        private ZipArchive? _templateArchive;
        private XmlReader? _templateReader;
        private bool _templateInSheetData;
        private int? _templateRowAhead;       // Row number of the template <row> the reader is parked on.
        private int _templateLastRow;         // For template rows written without an r attribute.
        private int _firstWrittenRow;         // First row written by the caller (0 = none).

        /// <summary>The template part this sheet fills, or null for a sheet written from scratch.</summary>
        internal string? TemplatePart { get; private set; }

        /// <summary>True when this sheet fills a sheet of a template workbook.</summary>
        public bool IsTemplateSheet => _templateSource != null;

        /// <summary>
        /// Number of template rows holding values or formulas that were replaced because a row with the same
        /// number was written (formatted but empty rows do not count).
        /// Lets callers detect a template footer that the data overran.
        /// </summary>
        public int ReplacedTemplateRows { get; private set; }

        /// <summary>
        /// First and last rows written by the caller, or null when none were written. Template ranges (tables,
        /// filters, validations, conditional formats) are stretched over this band. An edited sheet rewrites
        /// every row, so there the band is only the rows appended below the original last row: a table that
        /// ended on the last row grows over them, and nothing else moves.
        /// </summary>
        internal (int First, int Last)? WrittenBand
        {
            get
            {
                if (_editMode) return _editedLastRow > 0 && CurrentRow > _editedLastRow ? (_editedLastRow + 1, CurrentRow) : ((int, int)?)null;
                return _firstWrittenRow == 0 ? ((int, int)?)null : (_firstWrittenRow, CurrentRow);
            }
        }

        internal void AttachTemplate(TemplateSource source, string part)
        {
            _templateSource = source;
            TemplatePart = part;
            _templateArchive = source.OpenArchive();
            _templateReader = source.OpenXml(_templateArchive, part)
                ?? throw new InvalidOperationException($"The template has no part '{part}' for sheet '{Name}'.");
        }

        /// <summary>
        /// Returns the style of a cell in the template (for example the formatted first data row), so new
        /// rows can look exactly like the designer intended. Returns the default style when the cell has none.
        /// </summary>
        /// <example><code>var money = sheet.TemplateStyle("D5");</code></example>
        public CellStyle TemplateStyle(string cell)
        {
            if (_templateSource == null) throw new InvalidOperationException("TemplateStyle is only available on sheets opened from a template.");
            var (row, column) = RangeReference.ParseCell(cell);
            var wanted = CellReference.ToA1(row, column);
            using (var archive = _templateSource.OpenArchive())
            using (var r = _templateSource.OpenXml(archive, TemplatePart!)!)
            {
                while (r.Read())
                {
                    if (r.NodeType != XmlNodeType.Element) continue;
                    if (r.LocalName == "row")
                    {
                        // Rows are in ascending order; once past the wanted row, use the row's own style if any.
                        var rr = TemplateSource.Int(r.GetAttribute("r"));
                        if (rr.HasValue && rr.Value > row) break;
                        continue;
                    }
                    if (r.LocalName == "c" && string.Equals(r.GetAttribute("r"), wanted, StringComparison.OrdinalIgnoreCase))
                    {
                        var s = TemplateSource.Int(r.GetAttribute("s")) ?? 0;
                        return s == 0 ? Workbook.Styles.Default : new CellStyle(s);
                    }
                }
            }
            return Workbook.Styles.Default;
        }

        /// <summary>
        /// Copies the template sheet's start: the root element and everything before <c>sheetData</c>
        /// (sheet properties, views, column widths), dropping the stale <c>dimension</c> hint.
        /// </summary>
        private void CopyTemplateHeader()
        {
            var r = _templateReader!;
            _part.WriteStartDocument(true);
            XmlCopy.MoveToRoot(r);
            var rootEmpty = r.IsEmptyElement;
            _part.WriteStartElement(r.Prefix, r.LocalName, r.NamespaceURI);
            if (r.MoveToFirstAttribute())
            {
                do
                {
                    if (r.Prefix.Length == 0 && r.LocalName == "xmlns") continue; // Re-declared below.
                    _part.WriteAttributeString(r.Prefix, r.LocalName, r.NamespaceURI, r.Value);
                }
                while (r.MoveToNextAttribute());
                r.MoveToElement();
            }
            // Rows and cells are written without a prefix, so the main namespace must be the default one.
            _part.WriteAttributeString("xmlns", OpenXmlNamespaces.Main);
            if (rootEmpty) throw new System.IO.InvalidDataException($"Template sheet '{Name}' is empty.");
            r.Read();

            while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == 0))
            {
                if (r.NodeType != XmlNodeType.Element) { XmlCopy.CopyNode(r, _part); r.Read(); continue; }
                if (r.LocalName == "dimension") { r.Skip(); continue; }
                if (r.LocalName == "sheetData")
                {
                    var empty = r.IsEmptyElement;
                    _part.WriteStartElement(r.Prefix, r.LocalName, r.NamespaceURI);
                    r.Read();
                    _templateInSheetData = !empty;
                    if (_templateInSheetData) ParkOnNextTemplateRow();
                    // An edited sheet already holds every original row in its buffer and writes them from
                    // there, so the reader skips the whole sheetData instead of merging row by row.
                    if (_editMode) SkipAllTemplateRows();
                    return;
                }
                XmlCopy.CopySubtree(r, _part);
            }
            // A worksheet without sheetData is malformed; add one so the rows have somewhere to go.
            _part.WriteStartElement("sheetData", OpenXmlNamespaces.Main);
        }

        /// <summary>Moves the template reader to the next row of sheetData, or out of sheetData when there is none.</summary>
        private void ParkOnNextTemplateRow()
        {
            var r = _templateReader!;
            _templateRowAhead = null;
            while (_templateInSheetData)
            {
                if (r.NodeType == XmlNodeType.Element && r.LocalName == "row")
                {
                    _templateRowAhead = TemplateSource.Int(r.GetAttribute("r")) ?? _templateLastRow + 1;
                    return;
                }
                if (r.NodeType == XmlNodeType.EndElement && r.LocalName == "sheetData")
                {
                    r.Read();
                    _templateInSheetData = false;
                    return;
                }
                if (!r.Read()) { _templateInSheetData = false; return; }
            }
        }

        /// <summary>Skips every remaining template row; afterwards the reader is parked after sheetData.</summary>
        private void SkipAllTemplateRows()
        {
            var r = _templateReader!;
            while (_templateRowAhead.HasValue)
            {
                r.Skip();
                ParkOnNextTemplateRow();
            }
        }

        /// <summary>
        /// Before the caller starts row <paramref name="index"/>: copies template rows above it and drops the
        /// template row with the same number (the caller's row replaces it).
        /// </summary>
        private void MergeTemplateRowsBefore(int index)
        {
            if (_firstWrittenRow == 0) _firstWrittenRow = index;
            while (_templateRowAhead.HasValue && _templateRowAhead.Value <= index)
            {
                var r = _templateReader!;
                _templateLastRow = _templateRowAhead.Value;
                if (_templateRowAhead.Value < index)
                {
                    // Copied as XML without looking at its cells, so formulas must not treat it as known.
                    RecordUnknownRow(_templateRowAhead.Value);
                    XmlCopy.CopySubtree(r, _part);
                }
                else if (SkipTemplateRow(r)) ReplacedTemplateRows++;
                ParkOnNextTemplateRow();
            }
        }

        /// <summary>
        /// Skips the template row the reader is on and reports whether it held any value or formula. Formatted
        /// but empty rows (the styled first data row of a typical template) are expected to be replaced.
        /// </summary>
        private static bool SkipTemplateRow(XmlReader r)
        {
            if (r.IsEmptyElement)
            {
                r.Read();
                return false;
            }
            var hasContent = false;
            using (var sub = r.ReadSubtree())
            {
                while (sub.Read())
                    if (sub.NodeType == XmlNodeType.Element && (sub.LocalName == "v" || sub.LocalName == "f" || sub.LocalName == "is"))
                        hasContent = true;
            }
            r.Read(); // ReadSubtree leaves the parent on the row's end tag.
            return hasContent;
        }

        /// <summary>
        /// After the caller's last row: copies the remaining template rows below it, closes sheetData, then
        /// copies the rest of the template sheet (merges, rules, validations, drawings, tables), extending
        /// filters, conditional formats and validations that cover the first data row down to the last one.
        /// </summary>
        private void FinishTemplateSheet()
        {
            if (_tables.Count + _pictures.Count + _charts.Count + _comments.Count + _validations.Count + _conditionalFormats.Count + _hyperlinks.Count + _merges.Count > 0
                || _protection != null || _autoFilterRelative != null)
                throw new NotSupportedException($"Sheet '{Name}' fills a template, so its tables, rules, validations, links, comments, merges and protection come from the template. Add them to the template in Excel instead.");

            var r = _templateReader!;
            while (_templateRowAhead.HasValue)
            {
                _templateLastRow = _templateRowAhead.Value;
                if (_templateRowAhead.Value > CurrentRow) XmlCopy.CopySubtree(r, _part);
                else r.Skip(); // Unreachable in practice: rows at or above CurrentRow were merged already.
                ParkOnNextTemplateRow();
            }
            _part.WriteEndElement(); // sheetData

            var band = WrittenBand;
            XmlCopy.AttributeRewriter? extend = band == null ? null : (element, attribute, value) => ExtendTemplateRange(element, attribute, value, band.Value);
            while (!r.EOF && !(r.NodeType == XmlNodeType.EndElement && r.Depth == 0))
            {
                if (r.NodeType == XmlNodeType.Element) XmlCopy.CopySubtree(r, _part, extend);
                else { XmlCopy.CopyNode(r, _part); r.Read(); }
            }
            _part.WriteEndElement(); // worksheet
            _part.WriteEndDocument();
            _templateReader!.Dispose();
            _templateArchive!.Dispose();
        }

        /// <summary>
        /// Stretches a template range over the rows written by the caller. A range is extended when it covers
        /// the first written row (validations, conditional formats) or sits directly above it (a filter or
        /// table header), and ends before the last written row.
        /// </summary>
        internal static string? ExtendTemplateRange(string element, string attribute, string value, (int First, int Last) band)
        {
            var isFilter = element == "autoFilter" && attribute == "ref";
            var isSqref = (element == "conditionalFormatting" || element == "dataValidation") && attribute == "sqref";
            var isTable = element == "table" && attribute == "ref";
            if (!isFilter && !isSqref && !isTable) return value;

            var parts = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            var sb = new StringBuilder();
            foreach (var part in parts)
            {
                if (sb.Length > 0) sb.Append(' ');
                RangeReference range;
                try { range = RangeReference.Parse(part); }
                catch (ArgumentException) { sb.Append(part); continue; } // Whole columns or other forms: leave alone.
                var covers = isSqref
                    ? range.FromRow <= band.First && range.ToRow >= band.First
                    : range.FromRow < band.First && range.ToRow >= band.First - 1;
                if (covers && range.ToRow < band.Last)
                    range = new RangeReference(range.FromRow, range.FromColumn, band.Last, range.ToColumn);
                sb.Append(range.ToString());
            }
            return sb.ToString();
        }
    }
}
