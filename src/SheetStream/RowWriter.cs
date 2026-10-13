using System;
using System.Globalization;
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
    /// Every <c>WriteCell</c> overload returns the writer so calls can be chained:
    /// <c>sheet.Row().WriteCell("Total").WriteCell(1234.5, money)</c>. Cells are placed at consecutive
    /// columns starting at A; use <see cref="Skip(int)"/> to leave gaps.
    /// </remarks>
    public sealed class RowWriter
    {
        private readonly Worksheet _sheet;
        private readonly XmlWriter _w;
        private string _rowText = "";
        private CellStyle? _rowStyle;

        internal RowWriter(Worksheet sheet, XmlWriter writer)
        {
            _sheet = sheet;
            _w = writer;
        }

        /// <summary>1-based index of the row being written.</summary>
        public int RowIndex { get; private set; }

        /// <summary>1-based index the next cell will be written to.</summary>
        public int NextColumn { get; private set; }

        /// <summary>True between <see cref="Begin"/> and <see cref="End"/>.</summary>
        internal bool IsOpen { get; private set; }

        /// <summary>Opens the <c>&lt;row&gt;</c> element.</summary>
        internal void Begin(int rowIndex, CellStyle? style, double? height)
        {
            RowIndex = rowIndex;
            NextColumn = 1;
            _rowStyle = style != null && !style.IsDefault ? style : null;
            _rowText = rowIndex.ToString(CultureInfo.InvariantCulture);
            IsOpen = true;

            _w.WriteStartElement("row");
            _w.WriteAttributeString("r", _rowText);
            if (_rowStyle != null)
            {
                _w.WriteAttributeString("s", _rowStyle.Id.ToString(CultureInfo.InvariantCulture));
                _w.WriteAttributeString("customFormat", "1");
            }
            if (height.HasValue)
            {
                if (height.Value < 0 || height.Value > 409) throw new ArgumentOutOfRangeException(nameof(height), height, "Row height must be between 0 and 409 points.");
                _w.WriteAttributeString("ht", height.Value.ToString(CultureInfo.InvariantCulture));
                _w.WriteAttributeString("customHeight", "1");
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
            return _sheet.Workbook.Package.Options.UseSharedStrings ? WriteSharedString(value, style) : WriteInlineString(value, style);
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
            EndCell();
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
            EndCell();
            return this;
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
            EndCell();
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
            EndCell();
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
            EndCell();
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
            EndCell();
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
            EndCell();
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
            var effective = style ?? _rowStyle ?? (value.TimeOfDay == TimeSpan.Zero ? styles.DateStyle : styles.DateTimeStyle);
            BeginCell(effective, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value.ToOADate());
            _w.WriteEndElement();
            EndCell();
            return this;
        }

        /// <summary>Writes the local date-time of a <see cref="DateTimeOffset"/>; the offset itself is not stored (Excel has no time zones).</summary>
        public RowWriter WriteCell(DateTimeOffset value, CellStyle? style = null) => WriteCell(value.DateTime, style);

        /// <summary>Writes a duration as a fraction of days with the <c>[h]:mm:ss</c> format unless a style is given.</summary>
        public RowWriter WriteCell(TimeSpan value, CellStyle? style = null)
        {
            EnsureOpen();
            var effective = style ?? _rowStyle ?? _sheet.Workbook.Styles.DurationStyle;
            BeginCell(effective, null);
            _w.WriteStartElement("v");
            _w.WriteValue(value.TotalDays);
            _w.WriteEndElement();
            EndCell();
            return this;
        }

        /// <summary>
        /// Writes a value of runtime-determined type, dispatching to the typed overloads. Numbers, booleans,
        /// dates, time spans and strings are written natively; anything else is written as its invariant
        /// string representation. <c>null</c> and <see cref="DBNull"/> leave the column empty.
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

        /// <summary>
        /// Writes a formula cell, e.g. <c>"SUM(C2:C101)"</c>. A leading <c>=</c> is removed. No cached result
        /// is stored; the workbook asks Excel to recalculate on open, so the value appears when the file is
        /// opened in Excel (viewers that do not calculate will show the cell empty).
        /// </summary>
        public RowWriter WriteFormula(string formula, CellStyle? style = null)
        {
            if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("Formula cannot be empty.", nameof(formula));
            EnsureOpen();
            var text = formula.Trim();
            if (text[0] == '=') text = text.Substring(1);
            BeginCell(style, null);
            _w.WriteStartElement("f");
            _w.WriteString(XmlTextUtil.Sanitize(text));
            _w.WriteEndElement();
            EndCell();
            _sheet.Workbook.HasFormulas = true;
            return this;
        }

        private RowWriter WriteError(string errorCode, CellStyle? style)
        {
            BeginCell(style, "e");
            _w.WriteStartElement("v");
            _w.WriteString(errorCode);
            _w.WriteEndElement();
            EndCell();
            return this;
        }

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
            if (effective != null && !effective.IsDefault)
            {
                _w.WriteStartAttribute("s");
                _w.WriteValue(effective.Id);
                _w.WriteEndAttribute();
            }
            if (type != null) _w.WriteAttributeString("t", type);
        }

        private void EndCell()
        {
            _w.WriteEndElement();
            NextColumn++;
        }

        private void EnsureOpen()
        {
            if (!IsOpen)
                throw new InvalidOperationException("No row is open. Call Worksheet.Row() before writing cells.");
        }
    }
}
