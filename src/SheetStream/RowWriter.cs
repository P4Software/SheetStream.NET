using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Xml;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>
    /// Writes the cells of the current row, left to right. One instance exists per worksheet and is reused
    /// for every row, so writing a cell allocates nothing beyond what the XML writer buffers.
    /// </summary>
    /// <remarks>
    /// Every <c>Write*</c> method returns the writer so calls can be chained:
    /// <c>sheet.Row().WriteCell("Total").WriteCell(1234.5, money)</c>. Cells are placed at consecutive
    /// columns starting at A; use <see cref="Skip(int)"/> to leave gaps.
    /// </remarks>
    public sealed class RowWriter
    {
        private readonly Worksheet _sheet;
        private XmlWriter _w = null!;
        // Scratch space for formatting numbers without creating a string per cell (see WriteNumber).
        private readonly char[] _numberBuffer = new char[CharNumbers.BufferLength];
        private CellStyle? _rowStyle;
        private int _lastStyleId;

        // Cells of the common kinds are not handed to the XmlWriter one element and attribute at a time (its name
        // checks, state machine and virtual calls cost more than the cell's own text). Their finished XML is
        // composed here and passed over with a single WriteRaw. The buffer is a fixed size per sheet, flushed when
        // nearly full, so it never grows with the sheet; it is allocated when the first row opens and dropped when
        // the sheet is sealed. WriteRaw does no escaping or validation, so everything put here is escaped by hand
        // in exactly the way XmlWriter escapes text (see AppendText).
        internal const int RawCapacity = 1 << 14;

        // Room every composed cell needs on top of its text. The longest fixed part is an inline string with an
        // address, a style and xml:space (about 80 characters) and a number is at most 40, so 192 is generous.
        // Text is the only variable part and is written in pieces that flush as needed.
        private const int CellReserve = 192;
        private char[] _raw = Array.Empty<char>();
        private int _rawLength;

        // The row number as text, formatted once per row for the r attribute of any cell that needs one.
        private readonly char[] _rowDigits = new char[CharNumbers.BufferLength];
        private int _rowDigitsLength;

        // Column of the last <c> element written in the open row, 0 before the first. A cell is consecutive,
        // and may leave out its r attribute, exactly when NextColumn is this plus one.
        private int _lastWrittenColumn;

        // PackageOptions.WriteCellReferences, read once per row so each cell pays one bool test and no more.
        private bool _alwaysWriteReference;

        // True when the sheet wants each cell's display text (AutoFit, or a table header being captured).
        // Computed once per row so the common path pays a single bool test per cell.
        private bool _observe;

        // True when the sheet caches formula results and therefore wants every value recorded. Read once per
        // row, so with the feature off (the default) each cell pays one false branch and nothing else.
        private bool _cache;

        // Row style combined with a date / date-time / duration format, derived once per row on first use.
        private CellStyle? _rowDateStyle;
        private CellStyle? _rowDateTimeStyle;
        private CellStyle? _rowDurationStyle;

        internal RowWriter(Worksheet sheet)
        {
            _sheet = sheet;
        }

        /// <summary>1-based index of the row being written.</summary>
        public int RowIndex { get; private set; }

        /// <summary>1-based index the next cell will be written to.</summary>
        public int NextColumn { get; private set; }

        /// <summary>True between <see cref="Begin"/> and <see cref="End"/>.</summary>
        internal bool IsOpen { get; private set; }

        /// <summary>Connects the writer to the sheet body (the part itself, or the AutoFit spill file).</summary>
        internal void Attach(XmlWriter writer) => _w = writer;

        /// <summary>Re-reads whether the open row is observed, after a table was added on it.</summary>
        internal void RefreshObservation()
        {
            if (IsOpen) _observe = _sheet.ShouldObserveRow(RowIndex);
        }

        /// <summary>
        /// Opens the <c>&lt;row&gt;</c> element. <paramref name="rawAttributes"/> are attributes copied from an
        /// edited workbook's original row (height, hidden, outline, descent) and are written as they were.
        /// </summary>
        internal void Begin(int rowIndex, RowOptions? options, IReadOnlyList<RawAttribute>? rawAttributes = null)
        {
            RowIndex = rowIndex;
            NextColumn = 1;
            _lastWrittenColumn = 0;
            _alwaysWriteReference = _sheet.Workbook.Package.Options.WriteCellReferences;
            var style = options?.Style;
            _rowStyle = style != null && !style.IsDefault ? style : null;
            _rowDateStyle = _rowDateTimeStyle = _rowDurationStyle = null;
            _observe = _sheet.ShouldObserveRow(rowIndex);
            _cache = _sheet.IsCachingFormulaResults;
            IsOpen = true;

            if (_raw.Length == 0) _raw = new char[RawCapacity];
            FlushRaw(); // Nothing is pending after End; this keeps the ordering obvious.
            _rowDigitsLength = CharNumbers.Format(rowIndex, _rowDigits);
            _w.WriteStartElement("row");
            _w.WriteStartAttribute("r");
            WriteInteger(rowIndex);
            _w.WriteEndAttribute();
            if (rawAttributes != null)
                for (var i = 0; i < rawAttributes.Count; i++) rawAttributes[i].WriteTo(_w);
            if (_rowStyle != null)
            {
                _w.WriteAttributeString("s", _rowStyle.Id.ToString(CultureInfo.InvariantCulture));
                _w.WriteAttributeString("customFormat", "1");
            }
            if (options != null)
            {
                if (options.Height.HasValue)
                {
                    var h = options.Height.Value;
                    if (h < 0 || h > 409) throw new ArgumentOutOfRangeException(nameof(options), h, "Row height must be between 0 and 409 points.");
                    _w.WriteAttributeString("ht", h.ToString(CultureInfo.InvariantCulture));
                    _w.WriteAttributeString("customHeight", "1");
                }
                if (options.Hidden) _w.WriteAttributeString("hidden", "1");
                if (options.OutlineLevel > 0) _w.WriteAttributeString("outlineLevel", options.OutlineLevel.ToString(CultureInfo.InvariantCulture));
                if (options.Collapsed) _w.WriteAttributeString("collapsed", "1");
            }
        }

        /// <summary>Closes the <c>&lt;row&gt;</c> element if one is open.</summary>
        internal void End()
        {
            if (!IsOpen) return;
            // The row's cells are still in the staging buffer; they must reach the writer before its end tag.
            FlushRaw();
            _w.WriteEndElement();
            IsOpen = false;
        }

        /// <summary>
        /// Gives the staging buffer back once the sheet is sealed. Every write method checks for an open row
        /// first, so nothing can reach the empty buffer afterwards.
        /// </summary>
        internal void ReleaseBuffer() => _raw = Array.Empty<char>();

        /// <summary>Leaves <paramref name="count"/> columns empty.</summary>
        public RowWriter Skip(int count = 1)
        {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            EnsureOpen();
            NextColumn += count;
            return this;
        }

        /// <summary>Writes a cell that has a style but no value (useful for coloured filler cells).</summary>
        public RowWriter WriteEmpty(CellStyle? style = null)
        {
            EnsureOpen();
            var effective = style ?? _rowStyle;
            if (effective != null && !effective.IsDefault)
            {
                BeginCellRaw(effective, null);
                // XmlWriter closes a childless element as " />" (with the space); keep the file byte for byte.
                AppendFixed(" />");
            }
            NextColumn++;
            return this;
        }

        /// <summary>
        /// Writes a text cell. <c>null</c> leaves the column empty. By default the text goes through the
        /// shared string table; see <see cref="PackageOptions.UseSharedStrings"/>.
        /// </summary>
        public RowWriter WriteCell(string? value, CellStyle? style = null)
        {
            if (value == null) return WriteEmpty(style);
            return _sheet.Workbook.UseSharedStrings ? WriteSharedString(value, style) : WriteInlineString(value, style);
        }

        /// <summary>
        /// Writes a text cell inline in the sheet instead of through the shared string table. Use for
        /// values that are unlikely to repeat (free text, GUIDs) to keep the table small.
        /// </summary>
        public RowWriter WriteInlineString(string? value, CellStyle? style = null)
        {
            if (value == null) return WriteEmpty(style);
            EnsureOpen();
            BeginCellRaw(style, TypeInline);
            var text = XmlTextUtil.Sanitize(value);
            AppendFixed(XmlTextUtil.NeedsSpacePreserve(text) ? "><is><t xml:space=\"preserve\">" : "><is><t>");
            AppendText(text);
            EnsureRoom();
            AppendFixed("</t></is></c>");
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(value);
            return this;
        }

        private RowWriter WriteSharedString(string value, CellStyle? style)
        {
            EnsureOpen();
            // A new string that no longer fits the table's memory budget is written inline instead; strings
            // already in the table keep being shared.
            if (!_sheet.Workbook.SharedStrings.TryGetOrAdd(XmlTextUtil.Sanitize(value), out var index))
                return WriteInlineString(value, style);
            BeginCellRaw(style, TypeSharedString);
            AppendFixed("><v>");
            AppendInteger(index);
            AppendFixed("</v></c>");
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(value);
            return this;
        }

        /// <summary>
        /// Writes a rich-text cell made of differently formatted runs.
        /// </summary>
        /// <example>
        /// <code>row.WriteRichText("Total: ", new RichText("overdue") { Bold = true, Color = "#C00000" });</code>
        /// </example>
        public RowWriter WriteRichText(params RichText[] runs) => WriteRichText(null, runs);

        /// <summary>Writes a rich-text cell with a cell style (alignment, fill, borders) around the runs.</summary>
        public RowWriter WriteRichText(CellStyle? style, params RichText[] runs)
        {
            if (runs == null || runs.Length == 0) throw new ArgumentException("A rich-text cell needs at least one run.", nameof(runs));
            EnsureOpen();
            // Rich text is written inline: shared strings would need the same run markup and rarely repeat.
            BeginCell(style, "inlineStr");
            _w.WriteStartElement("is");
            foreach (var run in runs) (run ?? throw new ArgumentException("Runs cannot be null.", nameof(runs))).WriteTo(_w);
            _w.WriteEndElement();
            string? text = null;
            if (_observe)
            {
                var sb = new StringBuilder();
                foreach (var run in runs) sb.Append(run.Text);
                text = sb.ToString();
            }
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCell(text);
            return this;
        }

        /// <summary>
        /// Writes a text cell that links to a web page, file or e-mail address (<c>mailto:</c>). The built-in
        /// Hyperlink style is used unless another style is given.
        /// </summary>
        public RowWriter WriteHyperlink(string text, string url, CellStyle? style = null, string? tooltip = null)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentException("URL cannot be empty.", nameof(url));
            EnsureOpen();
            _sheet.RegisterHyperlink(RowIndex, NextColumn, url, null, tooltip);
            return WriteCell(text ?? url, style ?? _sheet.Workbook.Styles.Hyperlink);
        }

        /// <summary>Writes a text cell that jumps to a place in this workbook, e.g. <c>"'Summary'!A1"</c>.</summary>
        public RowWriter WriteInternalLink(string text, string location, CellStyle? style = null, string? tooltip = null)
        {
            if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("Location cannot be empty.", nameof(location));
            EnsureOpen();
            _sheet.RegisterHyperlink(RowIndex, NextColumn, null, location.TrimStart('#'), tooltip);
            return WriteCell(text ?? location, style ?? _sheet.Workbook.Styles.Hyperlink);
        }

        /// <summary>Writes a numeric cell. NaN and infinities become the <c>#NUM!</c> error, which is what Excel shows for them.</summary>
        public RowWriter WriteCell(double value, CellStyle? style = null)
        {
            EnsureOpen();
            if (double.IsNaN(value) || double.IsInfinity(value)) return WriteError("#NUM!", style);
            BeginCellRaw(style, null);
            AppendFixed("><v>");
            AppendNumber(value);
            AppendFixed("</v></c>");
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCellRaw(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell.</summary>
        public RowWriter WriteCell(float value, CellStyle? style = null) => WriteCell((double)value, style);

        /// <summary>Writes a numeric cell.</summary>
        public RowWriter WriteCell(int value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCellRaw(style, null);
            AppendFixed("><v>");
            AppendInteger(value);
            AppendFixed("</v></c>");
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCellRaw(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell. Values beyond 2^53 lose precision in Excel, which stores all numbers as doubles.</summary>
        public RowWriter WriteCell(long value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCellRaw(style, null);
            AppendFixed("><v>");
            AppendInteger(value);
            AppendFixed("</v></c>");
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCellRaw(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell. Excel keeps 15 significant digits.</summary>
        public RowWriter WriteCell(decimal value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCellRaw(style, null);
            AppendFixed("><v>");
            AppendNumber(value);
            AppendFixed("</v></c>");
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, (double)value);
            EndCellRaw(_observe ? NumberText((double)value) : null);
            return this;
        }

        /// <summary>Writes a boolean cell (shown as TRUE / FALSE).</summary>
        public RowWriter WriteCell(bool value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCellRaw(style, TypeBoolean);
            AppendFixed(value ? "><v>1</v></c>" : "><v>0</v></c>");
            if (_cache) _sheet.RecordText(RowIndex, NextColumn); // Booleans do not take part in SUM/AVERAGE over a range.
            EndCellRaw(value ? "TRUE" : "FALSE");
            return this;
        }

        /// <summary>
        /// Writes a date cell. Excel stores dates as serial numbers, so the value is written as
        /// Excel's serial for that moment and, when no style is given, formatted with a built-in date or
        /// date-time format depending on whether the value has a time component.
        /// </summary>
        public RowWriter WriteCell(DateTime value, CellStyle? style = null)
        {
            EnsureOpen();
            var styles = _sheet.Workbook.Styles;
            CellStyle effective;
            if (style != null) effective = style;
            else if (_rowStyle == null) effective = value.TimeOfDay == TimeSpan.Zero ? styles.DateStyle : styles.DateTimeStyle;
            // A row style (e.g. zebra fill) must not turn the date into a serial number: merge in a date format.
            else if (value.TimeOfDay == TimeSpan.Zero) effective = _rowDateStyle ??= styles.WithNumberFormat(_rowStyle, "mm-dd-yy");
            else effective = _rowDateTimeStyle ??= styles.WithNumberFormat(_rowStyle, "m/d/yy h:mm");
            // Not ToOADate: Excel's serials differ from OLE dates before 1 March 1900 (see ExcelDate).
            var serial = ExcelDate.ToSerial(value);
            BeginCellRaw(effective, null);
            AppendFixed("><v>");
            AppendNumber(serial);
            AppendFixed("</v></c>");
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, serial);
            // Display width follows the style's date format: "10/13/2026", "2026-10-13", "13-Oct-26"...
            EndCellRaw(_observe ? NumberText(serial) : null);
            return this;
        }

        /// <summary>Writes the local date-time of a <see cref="DateTimeOffset"/>; the offset itself is not stored (Excel has no time zones).</summary>
        public RowWriter WriteCell(DateTimeOffset value, CellStyle? style = null) => WriteCell(value.DateTime, style);

        /// <summary>Writes a duration as a fraction of days with the <c>[h]:mm:ss</c> format unless a style is given.</summary>
        public RowWriter WriteCell(TimeSpan value, CellStyle? style = null)
        {
            EnsureOpen();
            var styles = _sheet.Workbook.Styles;
            var effective = style ?? (_rowStyle == null ? styles.DurationStyle : _rowDurationStyle ??= styles.WithNumberFormat(_rowStyle, NumberFormats.Duration));
            BeginCellRaw(effective, null);
            AppendFixed("><v>");
            AppendNumber(value.TotalDays);
            AppendFixed("</v></c>");
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value.TotalDays);
            EndCellRaw(_observe ? NumberText(value.TotalDays) : null);
            return this;
        }

        /// <summary>
        /// Writes a value of runtime-determined type, dispatching to the typed overloads. Numbers, booleans,
        /// dates, time spans, strings and rich text are written natively; anything else is written as its
        /// invariant string representation. <c>null</c> and <see cref="DBNull"/> leave the column empty.
        /// </summary>
        public RowWriter WriteCell(object? value, CellStyle? style = null)
        {
            switch (value)
            {
                case null: return WriteEmpty(style);
                case DBNull _: return WriteEmpty(style);
                case string s: return WriteCell(s, style);
                case double d: return WriteCell(d, style);
                case float f: return WriteCell(f, style);
                case decimal m: return WriteCell(m, style);
                case int i: return WriteCell(i, style);
                case long l: return WriteCell(l, style);
                case short sh: return WriteCell((int)sh, style);
                case byte b: return WriteCell((int)b, style);
                case sbyte sb: return WriteCell((int)sb, style);
                case ushort us: return WriteCell((int)us, style);
                case uint ui: return WriteCell((long)ui, style);
                case ulong ul: return WriteCell((double)ul, style);
                case bool bo: return WriteCell(bo, style);
                case DateTime dt: return WriteCell(dt, style);
                case DateTimeOffset dto: return WriteCell(dto, style);
                case TimeSpan ts: return WriteCell(ts, style);
                case RichText rt: return WriteRichText(style, rt);
                case RichText[] rts: return WriteRichText(style, rts);
                case char c: return WriteCell(c.ToString(), style);
                case Guid g: return WriteCell(g.ToString(), style);
                case Enum e: return WriteCell(e.ToString(), style);
                // A style is not something a cell can hold. It gets here when it is passed among the values:
                // WriteCells(style, "a", "b") binds to the overload that takes only values. Writing the style's
                // type name into the sheet would be a silent mistake, so say what to write instead.
                case CellStyle _: throw new ArgumentException("A CellStyle was passed as a cell value. Use WriteCell(value, style) for one cell, or WriteCells(style, new[] { \"a\", \"b\" }) for several: the values go in a collection after the style.", nameof(value));
                case IFormattable fmt: return WriteCell(fmt.ToString(null, CultureInfo.InvariantCulture), style);
                default: return WriteCell(value.ToString(), style);
            }
        }

        /// <summary>Writes several cells from left to right using <see cref="WriteCell(object?, CellStyle?)"/>.</summary>
        public RowWriter WriteCells(params object?[] values)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            foreach (var value in values) WriteCell(value, null);
            return this;
        }

        /// <summary>Writes several cells from left to right, all with the same style.</summary>
        /// <example><code>sheet.Row(1).WriteCells(bold, new[] { "Id", "Customer", "Amount" });</code></example>
        /// <remarks>
        /// The values are a collection, not a list of arguments. <c>WriteCells(bold, "Id", "Customer")</c> binds
        /// to <see cref="WriteCells(object?[])"/> with the style as its first value, and throws to say so.
        /// </remarks>
        public RowWriter WriteCells(CellStyle? style, IEnumerable<object?> values)
        {
            if (values == null) throw new ArgumentNullException(nameof(values));
            foreach (var value in values) WriteCell(value, style);
            return this;
        }

        /// <summary>
        /// Writes a formula cell, e.g. <c>"SUM(C2:C101)"</c>. A leading <c>=</c> is removed. By default no cached
        /// result is stored; the workbook asks Excel to recalculate on open, so the value appears when the file
        /// is opened in Excel. With <see cref="Worksheet.CacheFormulaResults"/> enabled, simple aggregates over
        /// cells already written get their result stored too, so tools that do not calculate (viewers,
        /// ExcelReader, other libraries) show it. Pass the value yourself with the other overloads when you
        /// already know it.
        /// </summary>
        public RowWriter WriteFormula(string formula, CellStyle? style = null)
        {
            var text = PrepareFormula(formula);
            if (_cache && _sheet.TryEvaluateFormula(text, RowIndex, NextColumn, out var value))
                return WriteFormulaWithNumber(text, value, style);
            BeginCell(style, null);
            WriteF(text);
            // A later SUM over this cell cannot know its value, so it must not get one either.
            if (_cache) _sheet.RecordUnknown(RowIndex, NextColumn);
            EndCell(null);
            return this;
        }

        /// <summary>Writes a formula with a precomputed numeric result that non-calculating readers will see.</summary>
        public RowWriter WriteFormula(string formula, double cachedValue, CellStyle? style = null)
            => WriteFormulaWithNumber(PrepareFormula(formula), cachedValue, style);

        /// <summary>Shared by the explicit cached-value overload and the evaluator path; <paramref name="text"/> is already prepared.</summary>
        private RowWriter WriteFormulaWithNumber(string text, double cachedValue, CellStyle? style)
        {
            BeginCell(style, null);
            WriteF(text);
            _w.WriteStartElement("v");
            WriteNumber(cachedValue);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, cachedValue);
            EndCell(_observe ? NumberText(cachedValue) : null);
            return this;
        }

        /// <summary>Writes a formula with a precomputed text result.</summary>
        public RowWriter WriteFormula(string formula, string cachedValue, CellStyle? style = null)
        {
            if (cachedValue == null) throw new ArgumentNullException(nameof(cachedValue));
            var text = PrepareFormula(formula);
            BeginCell(style, "str");
            WriteF(text);
            _w.WriteElementString("v", XmlTextUtil.Sanitize(cachedValue));
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCell(cachedValue);
            return this;
        }

        /// <summary>Writes a formula with a precomputed boolean result.</summary>
        public RowWriter WriteFormula(string formula, bool cachedValue, CellStyle? style = null)
        {
            var text = PrepareFormula(formula);
            BeginCell(style, "b");
            WriteF(text);
            _w.WriteElementString("v", cachedValue ? "1" : "0");
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCell(cachedValue ? "TRUE" : "FALSE");
            return this;
        }

        private string PrepareFormula(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("Formula cannot be empty.", nameof(formula));
            EnsureOpen();
            _sheet.Workbook.HasFormulas = true;
            var text = formula.Trim();
            return text[0] == '=' ? text.Substring(1) : text;
        }

        private void WriteF(string text)
        {
            _w.WriteStartElement("f");
            _w.WriteString(XmlTextUtil.Sanitize(text));
            _w.WriteEndElement();
        }

        /// <summary>Writes an error cell (<c>#N/A</c>, <c>#DIV/0!</c>...), used when an edited workbook's error value is written back.</summary>
        internal RowWriter WriteErrorValue(string errorCode, CellStyle? style)
        {
            EnsureOpen();
            return WriteError(errorCode, style);
        }

        /// <summary>
        /// Writes a cell loaded from an edited workbook exactly as it was: the type attribute, any attributes the
        /// library does not model (<c>cm</c>, <c>vm</c>, <c>ph</c>) and the original inner XML (formula, cached
        /// value, rich inline text). Only the style is taken from <paramref name="style"/>, so a style change
        /// does not disturb the content. The caller records the cell's value for formula caching, since only it
        /// knows the typed value behind the raw XML.
        /// </summary>
        internal void WriteRawCell(string? type, IReadOnlyList<RawAttribute>? attributes, string innerXml, CellStyle? style)
        {
            EnsureOpen();
            BeginCell(style, type);
            if (attributes != null)
                for (var i = 0; i < attributes.Count; i++) attributes[i].WriteTo(_w);
            if (innerXml.Length > 0) _w.WriteRaw(innerXml);
            EndCell(null);
        }

        private RowWriter WriteError(string errorCode, CellStyle? style)
        {
            BeginCell(style, "e");
            _w.WriteStartElement("v");
            _w.WriteString(errorCode);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCell(errorCode);
            return this;
        }

        /// <summary>
        /// Approximates how a number will display under the cell's number format, for AutoFit: General text for
        /// unformatted numbers, the formatted digits plus literals (<c>$1,234.50</c>, <c>12.34%</c>) otherwise,
        /// and a value-independent stand-in for dates and times (<c>00/00/0000</c>). See <see cref="NumberDisplayFormat"/>.
        /// </summary>
        private string NumberText(double value) => _sheet.Workbook.Styles.DisplayFormat(_lastStyleId).Render(value);

        /// <summary>
        /// Opens a <c>&lt;c&gt;</c> element. The reference is only written when the cell is not the next one in
        /// its row, and then as two strings (cached column letters, then the row text computed once per row)
        /// so no per-cell string is allocated.
        /// </summary>
        private void BeginCell(CellStyle? style, string? type)
        {
            if (NextColumn > CellReference.MaxColumns)
                throw new InvalidOperationException("Row already has 16384 cells, which is Excel's column limit.");

            // Cells composed in the staging buffer come first in the file; this one goes through the XmlWriter.
            FlushRaw();
            _w.WriteStartElement("c");
            // The r attribute is optional (ECMA-376 18.3.1.4): a cell without one sits in the column after the
            // previous cell of its row, or in column A when it is the first. So it is written only when this
            // cell is NOT the next one in line (after Skip, an unstyled WriteEmpty, or any other gap), which
            // keeps the position exact while the common dense row pays nothing for it, in bytes or in time.
            // PackageOptions.WriteCellReferences switches the shortcut off for readers that need every address.
            if (_alwaysWriteReference || NextColumn != _lastWrittenColumn + 1)
            {
                _w.WriteStartAttribute("r");
                _w.WriteString(CellReference.ColumnLetters(NextColumn));
                WriteInteger(RowIndex);
                _w.WriteEndAttribute();
            }
            _lastWrittenColumn = NextColumn;

            var effective = style ?? _rowStyle;
            _lastStyleId = effective?.Id ?? 0;
            if (effective != null && !effective.IsDefault)
            {
                _w.WriteStartAttribute("s");
                WriteInteger(effective.Id);
                _w.WriteEndAttribute();
            }
            if (type != null) _w.WriteAttributeString("t", type);
        }

        /// <summary>
        /// Writes an integer as text content (or an attribute value) through the reusable buffer. Unlike
        /// <c>XmlWriter.WriteValue(int)</c> it creates no string, which is what keeps a cell allocation-free.
        /// </summary>
        private void WriteInteger(long value) => _w.WriteChars(_numberBuffer, 0, CharNumbers.Format(value, _numberBuffer));

        /// <summary>
        /// Writes a double like <c>XmlWriter.WriteValue(double)</c> does (shortest round-trip text, invariant
        /// culture), but through the reusable buffer. On <c>netstandard2.0</c>, which cannot format into a buffer,
        /// it falls back to the string-creating call.
        /// </summary>
        private void WriteNumber(double value)
        {
            var length = CharNumbers.Format(value, _numberBuffer);
            if (length < 0) _w.WriteValue(value);
            else _w.WriteChars(_numberBuffer, 0, length);
        }

        /// <summary>Decimal counterpart of <see cref="WriteNumber(double)"/>.</summary>
        private void WriteNumber(decimal value)
        {
            var length = CharNumbers.Format(value, _numberBuffer);
            if (length < 0) _w.WriteValue(value);
            else _w.WriteChars(_numberBuffer, 0, length);
        }

        private void EndCell(string? displayText)
        {
            _w.WriteEndElement();
            if (_observe && displayText != null) _sheet.Observe(RowIndex, NextColumn, displayText, _lastStyleId);
            NextColumn++;
        }

        // ---- Cells composed as raw XML --------------------------------------------------------------------------

        // Type attributes of the common cell kinds, with their leading space, as XmlWriter wrote them.
        private const string TypeSharedString = " t=\"s\"";
        private const string TypeInline = " t=\"inlineStr\"";
        private const string TypeBoolean = " t=\"b\"";

        /// <summary>
        /// Raw counterpart of <see cref="BeginCell"/>: appends <c>&lt;c</c>, the optional <c>r</c> and <c>s</c>
        /// attributes and <paramref name="typeAttribute"/> (already formatted, with its leading space) to the staging
        /// buffer, leaving the start tag open. The rules for when the address is written are the ones of
        /// <see cref="BeginCell"/>; keep the two in step.
        /// </summary>
        private void BeginCellRaw(CellStyle? style, string? typeAttribute)
        {
            if (NextColumn > CellReference.MaxColumns)
                throw new InvalidOperationException("Row already has 16384 cells, which is Excel's column limit.");

            // One check per cell guarantees room for everything fixed in it (see CellReserve).
            if (_raw.Length - _rawLength < CellReserve) FlushRaw();
            AppendFixed("<c");
            if (_alwaysWriteReference || NextColumn != _lastWrittenColumn + 1)
            {
                AppendFixed(" r=\"");
                AppendFixed(CellReference.ColumnLetters(NextColumn));
                AppendDigits(_rowDigits, _rowDigitsLength);
                AppendFixed("\"");
            }
            _lastWrittenColumn = NextColumn;

            var effective = style ?? _rowStyle;
            _lastStyleId = effective?.Id ?? 0;
            if (effective != null && !effective.IsDefault)
            {
                AppendFixed(" s=\"");
                AppendInteger(effective.Id);
                AppendFixed("\"");
            }
            if (typeAttribute != null) AppendFixed(typeAttribute);
        }

        /// <summary>Raw counterpart of <see cref="EndCell"/>: the closing tag is already in the buffer.</summary>
        private void EndCellRaw(string? displayText)
        {
            if (_observe && displayText != null) _sheet.Observe(RowIndex, NextColumn, displayText, _lastStyleId);
            NextColumn++;
        }

        /// <summary>Hands everything staged so far to the XmlWriter, in order. Does nothing when the buffer is empty.</summary>
        private void FlushRaw()
        {
            // An empty flush must not reach WriteRaw: it would close the open start tag, and a row with no cells
            // is written as a self-closing element.
            if (_rawLength == 0) return;
            var length = _rawLength;
            _rawLength = 0;
            _w.WriteRaw(_raw, 0, length);
        }

        /// <summary>Makes sure a short closing fragment fits after text that may have filled the buffer.</summary>
        private void EnsureRoom()
        {
            if (_raw.Length - _rawLength < CellReserve) FlushRaw();
        }

        /// <summary>Appends a short constant or cached string. The caller has reserved room (see <see cref="CellReserve"/>).</summary>
        private void AppendFixed(string text)
        {
            var raw = _raw;
            var position = _rawLength;
            for (var i = 0; i < text.Length; i++) raw[position + i] = text[i];
            _rawLength = position + text.Length;
        }

        private void AppendDigits(char[] digits, int length)
        {
            var raw = _raw;
            var position = _rawLength;
            for (var i = 0; i < length; i++) raw[position + i] = digits[i];
            _rawLength = position + length;
        }

        /// <summary>Appends an integer in plain digits, straight into the staging buffer.</summary>
        private void AppendInteger(long value) => _rawLength += CharNumbers.Format(value, _raw, _rawLength);

        /// <summary>
        /// Appends a double the way <c>XmlWriter.WriteValue(double)</c> writes it. Where the runtime cannot format
        /// into a buffer (<c>netstandard2.0</c>) the string-creating call is used, as before.
        /// </summary>
        private void AppendNumber(double value)
        {
            var length = CharNumbers.Format(value, _raw, _rawLength);
            if (length < 0) AppendFixed(XmlConvert.ToString(value));
            else _rawLength += length;
        }

        /// <summary>Decimal counterpart of <see cref="AppendNumber(double)"/>.</summary>
        private void AppendNumber(decimal value)
        {
            var length = CharNumbers.Format(value, _raw, _rawLength);
            if (length < 0) AppendFixed(XmlConvert.ToString(value));
            else _rawLength += length;
        }

        /// <summary>
        /// Appends element text the way <c>XmlWriter.WriteString</c> writes it: <c>&amp;</c>, <c>&lt;</c> and
        /// <c>&gt;</c> become entities and everything else, including quotes, tabs and line breaks (the writer is set
        /// to leave newlines alone), is copied unchanged. The text has been through
        /// <see cref="XmlTextUtil.Sanitize"/>, so it holds no character XML cannot carry. It is scanned once; runs
        /// without a special character are copied in bulk, and a string longer than the buffer goes out in pieces.
        /// </summary>
        private void AppendText(string text)
        {
            var runStart = 0;
            for (var i = 0; i < text.Length; i++)
            {
                var c = text[i];
                // '&', '<' and '>' are all at or below '>', so almost every character is rejected by one compare.
                if (c > '>') continue;
                string entity;
                if (c == '&') entity = "&amp;";
                else if (c == '<') entity = "&lt;";
                else if (c == '>') entity = "&gt;";
                else continue;

                AppendRun(text, runStart, i - runStart);
                if (_raw.Length - _rawLength < 8) FlushRaw();
                AppendFixed(entity);
                runStart = i + 1;
            }
            AppendRun(text, runStart, text.Length - runStart);
        }

        /// <summary>Copies <paramref name="count"/> characters of <paramref name="text"/>, flushing whenever the buffer fills.</summary>
        private void AppendRun(string text, int start, int count)
        {
            while (count > 0)
            {
                var room = _raw.Length - _rawLength;
                if (room == 0)
                {
                    FlushRaw();
                    room = _raw.Length;
                }
                var piece = count < room ? count : room;
                text.CopyTo(start, _raw, _rawLength, piece);
                _rawLength += piece;
                start += piece;
                count -= piece;
            }
        }

        private void EnsureOpen()
        {
            if (!IsOpen)
                throw new InvalidOperationException("No row is open. Call Worksheet.Row() before writing cells.");
        }
    }
}
