using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>
    /// A worksheet being streamed into the package. Rows are written forward-only through
    /// <see cref="Row()"/> / <see cref="Row(int)"/>; layout settings that the schema places before the
    /// sheet data (column widths, frozen panes) must be declared before the first row.
    /// </summary>
    /// <remarks>
    /// The worksheet XML is produced in three phases: a header (<c>sheetViews</c>, <c>cols</c>,
    /// <c>sheetData</c> start) written lazily when the first row begins; the rows; and a trailer
    /// (<c>autoFilter</c>, <c>mergeCells</c>) written when the sheet is sealed by adding the next sheet
    /// or saving the package. Not thread-safe.
    /// </remarks>
    public sealed class Worksheet
    {
        private readonly XmlWriter _writer;
        private readonly SortedDictionary<int, ColumnSettings> _columns = new SortedDictionary<int, ColumnSettings>();
        private readonly List<string> _merges = new List<string>();
        private readonly RowWriter _row;

        private int _frozenRows;
        private int _frozenColumns;
        private bool _headerWritten;
        private bool _sealed;

        internal Worksheet(Workbook workbook, int index, string name, XmlWriter writer)
        {
            Workbook = workbook;
            Index = index;
            Name = name;
            _writer = writer;
            _row = new RowWriter(this, writer);
        }

        /// <summary>The owning workbook.</summary>
        public Workbook Workbook { get; }

        /// <summary>1-based position in the workbook, also used as the sheet's part number.</summary>
        public int Index { get; }

        /// <summary>Tab name.</summary>
        public string Name { get; }

        /// <summary>Index of the last row started, or 0 if none.</summary>
        public int CurrentRow { get; private set; }

        /// <summary>True once the sheet has been closed; no further rows can be written.</summary>
        public bool IsSealed => _sealed;

        /// <summary>Range covered by the AutoFilter, as an absolute reference, or <c>null</c>.</summary>
        internal string? AutoFilterRange { get; private set; }

        private string? _autoFilterRelative;

        internal XmlWriter Writer => _writer;

        /// <summary>
        /// Returns the settings of a 1-based column. Must be called before the first row is written.
        /// </summary>
        /// <example><code>sheet.Column(1).SetWidth(15);</code></example>
        public ColumnSettings Column(int index)
        {
            if (index < 1 || index > CellReference.MaxColumns) throw new ArgumentOutOfRangeException(nameof(index), index, "Column index must be between 1 and 16384.");
            EnsureHeaderNotWritten("Column settings");
            return GetOrCreateColumn(index);
        }

        /// <summary>
        /// Returns a handle that applies settings to every column from <paramref name="from"/> to
        /// <paramref name="to"/> inclusive. Adjacent columns with identical settings are written as a single
        /// <c>&lt;col min max&gt;</c> range, so configuring thousands of columns costs a handful of bytes.
        /// Must be called before the first row is written.
        /// </summary>
        /// <example><code>sheet.Columns(1, 2000).SetWidth(12);</code></example>
        public ColumnRangeSettings Columns(int from, int to)
        {
            if (from < 1 || from > CellReference.MaxColumns) throw new ArgumentOutOfRangeException(nameof(from), from, "Column index must be between 1 and 16384.");
            if (to < from || to > CellReference.MaxColumns) throw new ArgumentOutOfRangeException(nameof(to), to, "Last column must be >= first column and <= 16384.");
            EnsureHeaderNotWritten("Column settings");
            return new ColumnRangeSettings(this, from, to);
        }

        /// <summary>
        /// Freezes the top <paramref name="rows"/> rows and the leftmost <paramref name="columns"/> columns
        /// so they stay visible while scrolling. Must be called before the first row is written.
        /// </summary>
        public Worksheet FreezePanes(int rows, int columns = 0)
        {
            if (rows < 0) throw new ArgumentOutOfRangeException(nameof(rows));
            if (columns < 0) throw new ArgumentOutOfRangeException(nameof(columns));
            EnsureHeaderNotWritten("Frozen panes");
            _frozenRows = rows;
            _frozenColumns = columns;
            return this;
        }

        /// <summary>
        /// Adds filter drop-downs to a header row spanning <paramref name="firstColumn"/> to
        /// <paramref name="lastColumn"/>. Excel extends the filter over the contiguous data below it.
        /// Can be called at any time before the sheet is sealed. Only one filter per sheet.
        /// </summary>
        public Worksheet AutoFilter(int headerRow, int firstColumn, int lastColumn)
        {
            ValidateRow(headerRow, nameof(headerRow));
            if (firstColumn < 1 || lastColumn < firstColumn || lastColumn > CellReference.MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(lastColumn), "Columns must satisfy 1 <= firstColumn <= lastColumn <= 16384.");
            EnsureNotSealed();
            _autoFilterRelative = CellReference.ToRange(headerRow, firstColumn, headerRow, lastColumn);
            AutoFilterRange = CellReference.ToAbsoluteRange(headerRow, firstColumn, headerRow, lastColumn);
            return this;
        }

        /// <summary>
        /// Merges the rectangle from (<paramref name="fromRow"/>, <paramref name="fromColumn"/>) to
        /// (<paramref name="toRow"/>, <paramref name="toColumn"/>). Only the top-left cell's value is shown.
        /// Can be called at any time before the sheet is sealed; overlapping merges are not checked.
        /// </summary>
        public Worksheet MergeCells(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            ValidateRow(fromRow, nameof(fromRow));
            ValidateRow(toRow, nameof(toRow));
            if (fromColumn < 1 || toColumn > CellReference.MaxColumns || toColumn < fromColumn || toRow < fromRow)
                throw new ArgumentException("Merge range must be a non-empty rectangle within the sheet.");
            if (fromRow == toRow && fromColumn == toColumn) return this; // A single cell is not a merge.
            EnsureNotSealed();
            _merges.Add(CellReference.ToRange(fromRow, fromColumn, toRow, toColumn));
            return this;
        }

        /// <summary>Starts the next row (current row + 1) and returns the writer for its cells.</summary>
        public RowWriter Row() => Row(CurrentRow + 1, null, null);

        /// <summary>
        /// Returns the writer for row <paramref name="index"/>. Returns the open row if it is the current one,
        /// starts a new row if the index is beyond it (gaps are allowed and simply left empty), and throws if
        /// the row has already been passed, because the sheet is streamed forward-only.
        /// </summary>
        public RowWriter Row(int index) => Row(index, null, null);

        /// <summary>
        /// Starts row <paramref name="index"/> with a default style and/or custom height. The style applies
        /// to cells written without their own style and to the row's empty cells.
        /// </summary>
        /// <param name="index">1-based row index; must be the current row or later.</param>
        /// <param name="style">Default style for the row, or <c>null</c>.</param>
        /// <param name="height">Row height in points, or <c>null</c> for Excel's default.</param>
        public RowWriter Row(int index, CellStyle? style, double? height = null)
        {
            ValidateRow(index, nameof(index));
            EnsureNotSealed();

            if (index == CurrentRow && _row.IsOpen)
            {
                if (style != null || height.HasValue)
                    throw new InvalidOperationException($"Row {index} is already open; its style and height cannot be changed after the first cell is written.");
                return _row;
            }
            if (index <= CurrentRow)
                throw new InvalidOperationException($"Row {index} has already been written. SheetStream streams rows forward-only; write rows in ascending order.");

            EnsureHeaderWritten();
            _row.End();
            CurrentRow = index;
            _row.Begin(index, style, height);
            return _row;
        }

        /// <summary>Closes the sheet: ends the open row, writes the trailer and releases the zip entry.</summary>
        internal void Seal()
        {
            if (_sealed) return;
            EnsureHeaderWritten(); // A sheet with no rows still needs a (possibly empty) sheetData element.
            _row.End();

            _writer.WriteEndElement(); // sheetData

            if (_autoFilterRelative != null)
            {
                _writer.WriteStartElement("autoFilter");
                _writer.WriteAttributeString("ref", _autoFilterRelative);
                _writer.WriteEndElement();
            }

            if (_merges.Count > 0)
            {
                _writer.WriteStartElement("mergeCells");
                _writer.WriteAttributeString("count", _merges.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var range in _merges)
                {
                    _writer.WriteStartElement("mergeCell");
                    _writer.WriteAttributeString("ref", range);
                    _writer.WriteEndElement();
                }
                _writer.WriteEndElement();
            }

            _writer.WriteEndElement(); // worksheet
            _writer.WriteEndDocument();
            _writer.Dispose();
            _sealed = true;
        }

        /// <summary>
        /// Writes everything that precedes the rows. CT_Worksheet is a strict sequence, so the order here
        /// (sheetViews, cols, sheetData) is not negotiable.
        /// </summary>
        private void EnsureHeaderWritten()
        {
            if (_headerWritten) return;
            _headerWritten = true;

            _writer.WriteStartDocument(true);
            _writer.WriteStartElement("worksheet", OpenXmlNamespaces.Main);
            _writer.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            if (_frozenRows > 0 || _frozenColumns > 0) WriteFrozenPanes();

            WriteColumnDefinitions();

            _writer.WriteStartElement("sheetData");
        }

        /// <summary>
        /// Writes <c>&lt;cols&gt;</c>, merging runs of adjacent columns that share the same width, visibility and
        /// style into one <c>&lt;col min max&gt;</c>. A sheet with 5,000 identically sized columns therefore
        /// emits a single element instead of 5,000, which matters for the wide exports this library targets.
        /// </summary>
        private void WriteColumnDefinitions()
        {
            ColumnSettings? runStart = null;
            var runEnd = 0;
            var hasCols = false;

            foreach (var column in _columns.Values)
            {
                if (!column.HasSettings) continue;
                if (runStart != null && column.Index == runEnd + 1 && column.HasSameSettingsAs(runStart))
                {
                    runEnd = column.Index;
                    continue;
                }
                if (runStart != null) WriteColumnRun(runStart, runEnd, ref hasCols);
                runStart = column;
                runEnd = column.Index;
            }
            if (runStart != null) WriteColumnRun(runStart, runEnd, ref hasCols);
            if (hasCols) _writer.WriteEndElement();
        }

        private void WriteColumnRun(ColumnSettings settings, int lastIndex, ref bool hasCols)
        {
            if (!hasCols)
            {
                _writer.WriteStartElement("cols");
                hasCols = true;
            }
            _writer.WriteStartElement("col");
            _writer.WriteAttributeString("min", settings.Index.ToString(CultureInfo.InvariantCulture));
            _writer.WriteAttributeString("max", lastIndex.ToString(CultureInfo.InvariantCulture));
            // The schema requires width; fall back to Excel's default when only a style or hidden flag is set.
            _writer.WriteAttributeString("width", (settings.Width ?? 8.43).ToString(CultureInfo.InvariantCulture));
            if (settings.Style != null && !settings.Style.IsDefault) _writer.WriteAttributeString("style", settings.Style.Id.ToString(CultureInfo.InvariantCulture));
            if (settings.IsHidden) _writer.WriteAttributeString("hidden", "1");
            if (settings.Width.HasValue) _writer.WriteAttributeString("customWidth", "1");
            _writer.WriteEndElement();
        }

        /// <summary>Returns (creating if needed) the settings object for a column; used by <see cref="ColumnRangeSettings"/>.</summary>
        internal ColumnSettings GetOrCreateColumn(int index)
        {
            if (!_columns.TryGetValue(index, out var settings))
            {
                settings = new ColumnSettings(index);
                _columns[index] = settings;
            }
            return settings;
        }

        private void WriteFrozenPanes()
        {
            // The active pane is the one below/right of the split; selection must name the same pane.
            var activePane = _frozenRows > 0 && _frozenColumns > 0 ? "bottomRight" : _frozenRows > 0 ? "bottomLeft" : "topRight";

            _writer.WriteStartElement("sheetViews");
            _writer.WriteStartElement("sheetView");
            _writer.WriteAttributeString("workbookViewId", "0");
            _writer.WriteStartElement("pane");
            if (_frozenColumns > 0) _writer.WriteAttributeString("xSplit", _frozenColumns.ToString(CultureInfo.InvariantCulture));
            if (_frozenRows > 0) _writer.WriteAttributeString("ySplit", _frozenRows.ToString(CultureInfo.InvariantCulture));
            _writer.WriteAttributeString("topLeftCell", CellReference.ToA1(_frozenRows + 1, _frozenColumns + 1));
            _writer.WriteAttributeString("activePane", activePane);
            _writer.WriteAttributeString("state", "frozen");
            _writer.WriteEndElement();
            _writer.WriteStartElement("selection");
            _writer.WriteAttributeString("pane", activePane);
            _writer.WriteEndElement();
            _writer.WriteEndElement();
            _writer.WriteEndElement();
        }

        private void EnsureHeaderNotWritten(string what)
        {
            EnsureNotSealed();
            if (_headerWritten)
                throw new InvalidOperationException($"{what} must be defined before the first row is written. SheetStream streams the sheet forward-only and the OpenXML schema places this element before the sheet data.");
        }

        private void EnsureNotSealed()
        {
            if (_sealed)
                throw new InvalidOperationException($"Worksheet '{Name}' has been sealed because another sheet was added or the package was saved. Finish writing a sheet before adding the next one.");
        }

        private static void ValidateRow(int row, string parameterName)
        {
            if (row < 1 || row > CellReference.MaxRows) throw new ArgumentOutOfRangeException(parameterName, row, "Row index must be between 1 and 1048576.");
        }
    }
}
