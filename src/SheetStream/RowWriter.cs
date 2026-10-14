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
        private string _rowText = "";
        private CellStyle? _rowStyle;
        private int _lastStyleId;

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
            var style = options?.Style;
            _rowStyle = style != null && !style.IsDefault ? style : null;
            _rowDateStyle = _rowDateTimeStyle = _rowDurationStyle = null;
            _rowText = rowIndex.ToString(CultureInfo.InvariantCulture);
            _observe = _sheet.ShouldObserveRow(rowIndex);
            _cache = _sheet.IsCachingFormulaResults;
            IsOpen = true;

            _w.WriteStartElement("row");
            _w.WriteAttributeString("r", _rowText);
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
            _w.WriteEndElement();
            IsOpen = false;
        }

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
                BeginCell(effective, null);
                _w.WriteEndElement();
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
            BeginCell(style, "inlineStr");
            _w.WriteStartElement("is");
            _w.WriteStartElement("t");
            var text = XmlTextUtil.Sanitize(value);
            if (XmlTextUtil.NeedsSpacePreserve(text)) _w.WriteAttributeString("xml", "space", null, "preserve");
            _w.WriteString(text);
            _w.WriteEndElement();
            _w.WriteEndElement();
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCell(value);
            return this;
        }

        private RowWriter WriteSharedString(string value, CellStyle? style)
        {
            EnsureOpen();
            var index = _sheet.Workbook.SharedStrings.GetOrAdd(XmlTextUtil.Sanitize(value));
            BeginCell(style, "s");
            _w.WriteStartElement("v");
            _w.WriteValue(index);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCell(value);
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
            BeginCell(style, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCell(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell.</summary>
        public RowWriter WriteCell(float value, CellStyle? style = null) => WriteCell((double)value, style);

        /// <summary>Writes a numeric cell.</summary>
        public RowWriter WriteCell(int value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCell(style, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCell(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell. Values beyond 2^53 lose precision in Excel, which stores all numbers as doubles.</summary>
        public RowWriter WriteCell(long value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCell(style, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCell(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell. Excel keeps 15 significant digits.</summary>
        public RowWriter WriteCell(decimal value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCell(style, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, (double)value);
            EndCell(_observe ? NumberText((double)value) : null);
            return this;
        }

        /// <summary>Writes a boolean cell (shown as TRUE / FALSE).</summary>
        public RowWriter WriteCell(bool value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCell(style, "b");
            _w.WriteStartElement("v");
            _w.WriteString(value ? "1" : "0");
            _w.WriteEndElement();
            if (_cache) _sheet.RecordText(RowIndex, NextColumn); // Booleans do not take part in SUM/AVERAGE over a range.
            EndCell(value ? "TRUE" : "FALSE");
            return this;
        }

        /// <summary>
        /// Writes a date cell. Excel stores dates as serial numbers, so the value is written as
        /// <see cref="DateTime.ToOADate"/> and, when no style is given, formatted with a built-in date or
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
            BeginCell(effective, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value.ToOADate());
            _w.WriteEndElement();
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value.ToOADate());
            // Display width follows the style's date format: "10/13/2026", "2026-10-13", "13-Oct-26"...
            EndCell(_observe ? NumberText(value.ToOADate()) : null);
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
            BeginCell(effective, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value.TotalDays);
            _w.WriteEndElement();
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value.TotalDays);
            EndCell(_observe ? NumberText(value.TotalDays) : null);
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
            _w.WriteValue(cachedValue);
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
        /// Opens a <c>&lt;c&gt;</c> element. The reference is written as two strings (cached column letters,
        /// then the row text computed once per row) so no per-cell string is allocated.
        /// </summary>
        private void BeginCell(CellStyle? style, string? type)
        {
            if (NextColumn > CellReference.MaxColumns)
                throw new InvalidOperationException("Row already has 16384 cells, which is Excel's column limit.");

            _w.WriteStartElement("c");
            _w.WriteStartAttribute("r");
            _w.WriteString(CellReference.ColumnLetters(NextColumn));
            _w.WriteString(_rowText);
            _w.WriteEndAttribute();

            var effective = style ?? _rowStyle;
            _lastStyleId = effective?.Id ?? 0;
            if (effective != null && !effective.IsDefault)
            {
                _w.WriteStartAttribute("s");
                _w.WriteValue(effective.Id);
                _w.WriteEndAttribute();
            }
            if (type != null) _w.WriteAttributeString("t", type);
        }

        private void EndCell(string? displayText)
        {
            _w.WriteEndElement();
            if (_observe && displayText != null) _sheet.Observe(RowIndex, NextColumn, displayText, _lastStyleId);
            NextColumn++;
        }

        private void EnsureOpen()
        {
            if (!IsOpen)
                throw new InvalidOperationException("No row is open. Call Worksheet.Row() before writing cells.");
        }
    }
}
