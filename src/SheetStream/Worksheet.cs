using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using SheetStream.Internal;
using SheetStream.Styles;

namespace SheetStream
{
    /// <summary>Formatting of a whole row, for <see cref="Worksheet.Row(int, RowOptions)"/>.</summary>
    public sealed class RowOptions
    {
        /// <summary>Default style for cells of the row written without their own style.</summary>
        public CellStyle? Style { get; set; }

        /// <summary>Row height in points (0 to 409), or null for Excel's default.</summary>
        public double? Height { get; set; }

        /// <summary>Hides the row.</summary>
        public bool Hidden { get; set; }

        /// <summary>Outline (grouping) level 0 to 7. Consecutive rows with a level form a collapsible group.</summary>
        public int OutlineLevel { get; set; }

        /// <summary>Marks the row as the collapsed summary row of a group.</summary>
        public bool Collapsed { get; set; }
    }

    /// <summary>
    /// A worksheet being streamed into the package. Rows are written forward-only through
    /// <see cref="Row()"/> / <see cref="Row(int)"/>; layout settings that the schema places before the
    /// sheet data (column widths, frozen panes, tab colour) must be declared before the first row, unless
    /// <see cref="AutoFitColumns"/> or random-access <see cref="Cells"/> are used, which defer the header.
    /// </summary>
    /// <remarks>
    /// The worksheet XML is produced in three phases: a header (<c>sheetPr</c>, <c>sheetViews</c>,
    /// <c>sheetFormatPr</c>, <c>cols</c>, <c>sheetData</c> start); the rows; and a trailer (protection,
    /// filters, merges, conditional formats, validations, hyperlinks, print settings, drawings, tables)
    /// written when the sheet is sealed by adding the next sheet or saving the package. Parts that live
    /// outside the sheet (tables, drawings, charts, comments, pivot tables, the sheet's relationships) are written right
    /// after it. Not thread-safe.
    /// </remarks>
    public sealed partial class Worksheet
    {
        private readonly XmlWriter _part;
        private readonly SortedDictionary<int, ColumnSettings> _columns = new SortedDictionary<int, ColumnSettings>();
        private readonly List<string> _merges = new List<string>();
        private readonly RowWriter _row;

        // Body: where <row> elements go. Normally the part itself; a temp file when AutoFit defers the header.
        private XmlWriter? _body;
        private FileStream? _spill;
        private bool _bodyStarted;
        private bool _headerWritten;
        private bool _sealed;

        // Header settings.
        private int _frozenRows;
        private int _frozenColumns;
        private string? _tabColor;
        private bool _showGridLines = true;
        private int? _zoom;
        private bool _rightToLeft;
        private int _maxRowOutlineLevel;

        private string? _autoFilterRelative;

        internal Worksheet(Workbook workbook, int index, string name, XmlWriter part)
        {
            Workbook = workbook;
            Index = index;
            Name = name;
            _part = part;
            _row = new RowWriter(this);
            PageSetup = new PageSetup(this);
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

        /// <summary>Tab visibility. Can be changed until the package is saved. At least one sheet must stay visible.</summary>
        public SheetVisibility Visibility { get; set; }

        /// <summary>Print settings (orientation, margins, header and footer, print titles...).</summary>
        public PageSetup PageSetup { get; }

        /// <summary>Range covered by the AutoFilter, as an absolute reference, or <c>null</c>.</summary>
        internal string? AutoFilterRange { get; private set; }

        /// <summary>
        /// Tab colour as hex ("#C00000"). Must be set before the first row; it lives in the sheet header.
        /// </summary>
        public string? TabColor
        {
            get => _tabColor;
            set
            {
                EnsureHeaderNotWritten("Tab colour");
                _tabColor = value == null ? null : XmlTextUtil.NormalizeColor(value, nameof(TabColor));
            }
        }

        /// <summary>Show the cell gridlines on screen. Default true. Must be set before the first row.</summary>
        public bool ShowGridLines
        {
            get => _showGridLines;
            set { EnsureHeaderNotWritten("Gridline visibility"); _showGridLines = value; }
        }

        /// <summary>Zoom percentage (10 to 400). Must be set before the first row.</summary>
        public int? Zoom
        {
            get => _zoom;
            set
            {
                if (value.HasValue && (value < 10 || value > 400)) throw new ArgumentOutOfRangeException(nameof(Zoom), value, "Zoom must be between 10 and 400 percent.");
                EnsureHeaderNotWritten("Zoom");
                _zoom = value;
            }
        }

        /// <summary>Lays the sheet out right to left (Arabic, Hebrew). Must be set before the first row.</summary>
        public bool RightToLeft
        {
            get => _rightToLeft;
            set { EnsureHeaderNotWritten("Right-to-left layout"); _rightToLeft = value; }
        }

        /// <summary>
        /// Returns the settings of a 1-based column. Must be called before the first row is written
        /// (or any time before saving when AutoFit or random-access cells are used).
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
        /// Can be called at any time before the sheet is sealed. Only one filter per sheet, and not on a
        /// range covered by a table (tables carry their own filter).
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

        /// <summary>Merges an A1 range such as "A1:D1".</summary>
        public Worksheet MergeCells(string range)
        {
            var r = RangeReference.Parse(range);
            return MergeCells(r.FromRow, r.FromColumn, r.ToRow, r.ToColumn);
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
            => Row(index, style == null && !height.HasValue ? null : new RowOptions { Style = style, Height = height });

        /// <summary>Starts row <paramref name="index"/> with full row options (style, height, hidden, outline level).</summary>
        public RowWriter Row(int index, RowOptions? options)
        {
            if (_buffer != null)
                throw new InvalidOperationException($"Worksheet '{Name}' is in random-access mode because Cells[...] was used or the sheet was opened with Worksheets.Edit. Write all of its cells through Cells, or use Row() on a sheet that never touches Cells.");
            _streamedRows = true;
            return RowCore(index, options, null);
        }

        /// <summary>
        /// Row start shared by streaming writes and the random-access buffer flush. <paramref name="rawAttributes"/>
        /// carries the original row attributes of an edited workbook (null everywhere else).
        /// </summary>
        private RowWriter RowCore(int index, RowOptions? options, IReadOnlyList<RawAttribute>? rawAttributes)
        {
            ValidateRow(index, nameof(index));
            EnsureNotSealed();

            if (index == CurrentRow && _row.IsOpen)
            {
                if (options != null)
                    throw new InvalidOperationException($"Row {index} is already open; its options cannot be changed after it was started.");
                return _row;
            }
            if (index <= CurrentRow)
                throw new InvalidOperationException($"Row {index} has already been written. SheetStream streams rows forward-only; write rows in ascending order, or use sheet.Cells[row, column] for random access.");

            EnsureBodyStarted();
            _row.End();
            if (_templateSource != null) MergeTemplateRowsBefore(index);
            else if (_firstWrittenRow == 0) _firstWrittenRow = index;
            CurrentRow = index;
            if (options != null)
            {
                if (options.OutlineLevel < 0 || options.OutlineLevel > 7) throw new ArgumentOutOfRangeException(nameof(options), options.OutlineLevel, "Outline level must be between 0 and 7.");
                if (options.OutlineLevel > _maxRowOutlineLevel) _maxRowOutlineLevel = options.OutlineLevel;
            }
            _row.Begin(index, options, rawAttributes);
            return _row;
        }

        /// <summary>Throws when the sheet can no longer change (used by feature builders).</summary>
        internal void EnsureWritable() => EnsureNotSealed();

        /// <summary>Closes the sheet: ends the open row, writes the trailer, the sheet's sub-parts and its relationships.</summary>
        internal void Seal()
        {
            if (_sealed) return;
            FlushBuffer();
            EnsureBodyStarted(); // A sheet with no rows still needs a (possibly empty) sheetData element.
            _row.End();

            if (_templateSource != null)
            {
                FinishTemplateSheet();
                _part.Dispose();
                _sealed = true;
                return;
            }

            if (_spill != null)
            {
                // AutoFit mode: the rows went to a temp file. Now that every width is known, write the header
                // and copy the rows in. Memory stays flat because the copy streams in 64K chunks.
                _body!.Flush();
                _body.Dispose();
                ApplyAutoFitWidths();
                EnsureHeaderWritten();
                _part.Flush();
                _spill.Position = 0;
                using (var reader = new StreamReader(_spill, new UTF8Encoding(false), false, 1 << 16, leaveOpen: true))
                {
                    var buffer = new char[1 << 15];
                    int n;
                    while ((n = reader.Read(buffer, 0, buffer.Length)) > 0) _part.WriteRaw(buffer, 0, n);
                }
                _spill.Dispose();
                _spill = null;
            }

            _part.WriteEndElement(); // sheetData
            var rels = new RelationshipSet();
            WriteTrailer(rels);
            _part.WriteEndElement(); // worksheet
            _part.WriteEndDocument();
            _part.Dispose();
            _sealed = true;

            WriteSubParts(rels);
        }

        /// <summary>
        /// Makes sure row output has somewhere to go. Streaming mode writes the header now and sends rows to
        /// the part; AutoFit mode sends rows to a temp file and defers the header until widths are known.
        /// </summary>
        private void EnsureBodyStarted()
        {
            if (_bodyStarted) return;
            _bodyStarted = true;
            if (_autoFit)
            {
                var path = Path.GetTempFileName();
                _spill = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 1 << 16, FileOptions.DeleteOnClose);
                _body = XmlWriter.Create(_spill, new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(false),
                    ConformanceLevel = ConformanceLevel.Fragment,
                    CloseOutput = false,
                    NewLineHandling = NewLineHandling.None,
                });
            }
            else
            {
                EnsureHeaderWritten();
                _body = _part;
            }
            _row.Attach(_body);
        }

        /// <summary>
        /// Writes everything that precedes the rows. CT_Worksheet is a strict sequence, so the order here
        /// (sheetPr, sheetViews, sheetFormatPr, cols, sheetData) is not negotiable.
        /// </summary>
        private void EnsureHeaderWritten()
        {
            if (_headerWritten) return;
            _headerWritten = true;
            if (_templateSource != null)
            {
                CopyTemplateHeader();
                return;
            }

            _part.WriteStartDocument(true);
            _part.WriteStartElement("worksheet", OpenXmlNamespaces.Main);
            _part.WriteAttributeString("xmlns", "r", null, OpenXmlNamespaces.Relationships);

            if (_tabColor != null || PageSetup.FitToPage)
            {
                _part.WriteStartElement("sheetPr");
                if (_tabColor != null)
                {
                    _part.WriteStartElement("tabColor");
                    _part.WriteAttributeString("rgb", _tabColor);
                    _part.WriteEndElement();
                }
                if (PageSetup.FitToPage)
                {
                    _part.WriteStartElement("pageSetUpPr");
                    _part.WriteAttributeString("fitToPage", "1");
                    _part.WriteEndElement();
                }
                _part.WriteEndElement();
            }

            WriteSheetViews();

            var maxColOutline = 0;
            foreach (var c in _columns.Values) if (c.OutlineLevel > maxColOutline) maxColOutline = c.OutlineLevel;
            _part.WriteStartElement("sheetFormatPr");
            _part.WriteAttributeString("defaultRowHeight", "15");
            // Excel sizes the outline gutter from these; the row level is only known here when the header is deferred.
            if (_maxRowOutlineLevel > 0) _part.WriteAttributeString("outlineLevelRow", _maxRowOutlineLevel.ToString(CultureInfo.InvariantCulture));
            if (maxColOutline > 0) _part.WriteAttributeString("outlineLevelCol", maxColOutline.ToString(CultureInfo.InvariantCulture));
            _part.WriteEndElement();

            WriteColumnDefinitions();

            _part.WriteStartElement("sheetData");
        }

        private void WriteSheetViews()
        {
            var frozen = _frozenRows > 0 || _frozenColumns > 0;
            if (!frozen && _showGridLines && !_zoom.HasValue && !_rightToLeft) return;

            _part.WriteStartElement("sheetViews");
            _part.WriteStartElement("sheetView");
            if (!_showGridLines) _part.WriteAttributeString("showGridLines", "0");
            if (_rightToLeft) _part.WriteAttributeString("rightToLeft", "1");
            if (_zoom.HasValue) _part.WriteAttributeString("zoomScale", _zoom.Value.ToString(CultureInfo.InvariantCulture));
            _part.WriteAttributeString("workbookViewId", "0");
            if (frozen)
            {
                // The active pane is the one below/right of the split; selection must name the same pane.
                var activePane = _frozenRows > 0 && _frozenColumns > 0 ? "bottomRight" : _frozenRows > 0 ? "bottomLeft" : "topRight";
                _part.WriteStartElement("pane");
                if (_frozenColumns > 0) _part.WriteAttributeString("xSplit", _frozenColumns.ToString(CultureInfo.InvariantCulture));
                if (_frozenRows > 0) _part.WriteAttributeString("ySplit", _frozenRows.ToString(CultureInfo.InvariantCulture));
                _part.WriteAttributeString("topLeftCell", CellReference.ToA1(_frozenRows + 1, _frozenColumns + 1));
                _part.WriteAttributeString("activePane", activePane);
                _part.WriteAttributeString("state", "frozen");
                _part.WriteEndElement();
                _part.WriteStartElement("selection");
                _part.WriteAttributeString("pane", activePane);
                _part.WriteEndElement();
            }
            _part.WriteEndElement();
            _part.WriteEndElement();
        }

        /// <summary>
        /// Writes <c>&lt;cols&gt;</c>, merging runs of adjacent columns that share the same width, visibility,
        /// outline level and style into one <c>&lt;col min max&gt;</c>. A sheet with 5,000 identically sized
        /// columns therefore emits a single element instead of 5,000, which matters for wide exports.
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
            if (hasCols) _part.WriteEndElement();
        }

        private void WriteColumnRun(ColumnSettings settings, int lastIndex, ref bool hasCols)
        {
            if (!hasCols)
            {
                _part.WriteStartElement("cols");
                hasCols = true;
            }
            _part.WriteStartElement("col");
            _part.WriteAttributeString("min", settings.Index.ToString(CultureInfo.InvariantCulture));
            _part.WriteAttributeString("max", lastIndex.ToString(CultureInfo.InvariantCulture));
            // The schema requires width; fall back to Excel's default when only a style, flag or level is set.
            _part.WriteAttributeString("width", Math.Round(settings.Width ?? 8.43, 2).ToString(CultureInfo.InvariantCulture));
            if (settings.Style != null && !settings.Style.IsDefault) _part.WriteAttributeString("style", settings.Style.Id.ToString(CultureInfo.InvariantCulture));
            if (settings.IsHidden) _part.WriteAttributeString("hidden", "1");
            if (settings.Width.HasValue) _part.WriteAttributeString("customWidth", "1");
            if (settings.OutlineLevel > 0) _part.WriteAttributeString("outlineLevel", settings.OutlineLevel.ToString(CultureInfo.InvariantCulture));
            _part.WriteEndElement();
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

        /// <summary>
        /// Writes the elements after sheetData in CT_Worksheet order: sheetProtection, autoFilter, mergeCells,
        /// conditionalFormatting, dataValidations, hyperlinks, printOptions, pageMargins, pageSetup,
        /// headerFooter, drawing, legacyDrawing, tableParts. Relationship ids are allocated as they are needed.
        /// </summary>
        private void WriteTrailer(RelationshipSet rels)
        {
            _protection?.WriteTo(_part);

            if (_autoFilterRelative != null)
            {
                foreach (var t in _tables)
                    if (t.Range.Overlaps(RangeReference.Parse(_autoFilterRelative)))
                        throw new InvalidOperationException($"Worksheet '{Name}' has an AutoFilter on {_autoFilterRelative} that overlaps table '{t.Name}'. Tables have their own filter; remove the sheet AutoFilter.");
                _part.WriteStartElement("autoFilter");
                _part.WriteAttributeString("ref", _autoFilterRelative);
                _part.WriteEndElement();
            }

            if (_merges.Count > 0)
            {
                _part.WriteStartElement("mergeCells");
                _part.WriteAttributeString("count", _merges.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var range in _merges)
                {
                    _part.WriteStartElement("mergeCell");
                    _part.WriteAttributeString("ref", range);
                    _part.WriteEndElement();
                }
                _part.WriteEndElement();
            }

            var priority = 1;
            foreach (var cf in _conditionalFormats) priority = cf.WriteTo(_part, priority);

            if (_validations.Count > 0)
            {
                _part.WriteStartElement("dataValidations");
                _part.WriteAttributeString("count", _validations.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var v in _validations) v.WriteTo(_part);
                _part.WriteEndElement();
            }

            if (_hyperlinks.Count > 0)
            {
                _part.WriteStartElement("hyperlinks");
                foreach (var h in _hyperlinks)
                {
                    _part.WriteStartElement("hyperlink");
                    _part.WriteAttributeString("ref", CellReference.ToA1(h.Row, h.Column));
                    if (h.Url != null) _part.WriteAttributeString("id", OpenXmlNamespaces.Relationships, rels.Add(OpenXmlNamespaces.RelHyperlink, h.Url, external: true));
                    if (h.Location != null) _part.WriteAttributeString("location", h.Location);
                    if (h.Tooltip != null) _part.WriteAttributeString("tooltip", XmlTextUtil.Sanitize(h.Tooltip));
                    _part.WriteEndElement();
                }
                _part.WriteEndElement();
            }

            PageSetup.WriteTo(_part);

            if (_pictures.Count > 0 || _charts.Count > 0)
            {
                _drawingNumber = ++Workbook.DrawingCount;
                _part.WriteStartElement("drawing");
                _part.WriteAttributeString("id", OpenXmlNamespaces.Relationships,
                    rels.Add(OpenXmlNamespaces.RelDrawing, "../drawings/drawing" + _drawingNumber.ToString(CultureInfo.InvariantCulture) + ".xml"));
                _part.WriteEndElement();
            }

            if (_comments.Count > 0)
            {
                _commentsNumber = ++Workbook.CommentsCount;
                var n = _commentsNumber.ToString(CultureInfo.InvariantCulture);
                _part.WriteStartElement("legacyDrawing");
                _part.WriteAttributeString("id", OpenXmlNamespaces.Relationships, rels.Add(OpenXmlNamespaces.RelVmlDrawing, "../drawings/vmlDrawing" + n + ".vml"));
                _part.WriteEndElement();
                rels.Add(OpenXmlNamespaces.RelComments, "../comments" + n + ".xml");
            }

            if (_tables.Count > 0)
            {
                _part.WriteStartElement("tableParts");
                _part.WriteAttributeString("count", _tables.Count.ToString(CultureInfo.InvariantCulture));
                foreach (var t in _tables)
                {
                    _part.WriteStartElement("tablePart");
                    _part.WriteAttributeString("id", OpenXmlNamespaces.Relationships, rels.Add(OpenXmlNamespaces.RelTable, "../tables/table" + t.Id.ToString(CultureInfo.InvariantCulture) + ".xml"));
                    _part.WriteEndElement();
                }
                _part.WriteEndElement();
            }
        }

        private int _drawingNumber;
        private int _commentsNumber;

        /// <summary>Writes the parts referenced from the sheet trailer, then the sheet's own <c>.rels</c>.</summary>
        private void WriteSubParts(RelationshipSet rels)
        {
            var package = Workbook.Package;
            foreach (var t in _tables)
            {
                var part = "xl/tables/table" + t.Id.ToString(CultureInfo.InvariantCulture) + ".xml";
                using (var w = package.OpenPart(part)) t.WriteTo(w);
                package.RegisterPart(part, OpenXmlNamespaces.CtTable);
            }
            if (_drawingNumber > 0) DrawingWriter.Write(package, Workbook, _drawingNumber, _pictures, _charts);
            if (_commentsNumber > 0) CommentsWriter.Write(package, _commentsNumber, Index, _comments, out _, out _);
            // A pivot table is referenced only from the sheet's .rels, never from the sheet XML, so its
            // relationship can be allocated here, after the sheet part is closed.
            foreach (var p in _pivotTables)
            {
                rels.Add(OpenXmlNamespaces.RelPivotTable, "../pivotTables/pivotTable" + p.Number.ToString(CultureInfo.InvariantCulture) + ".xml");
                PivotTableWriter.Write(package, Workbook, p);
            }
            if (rels.Count > 0)
                using (var w = package.OpenPart("xl/worksheets/_rels/sheet" + Index.ToString(CultureInfo.InvariantCulture) + ".xml.rels"))
                    rels.WriteTo(w);
        }

        /// <summary>Throws when a header-level setting is changed after the header was written.</summary>
        internal void EnsureHeaderNotWritten(string what)
        {
            EnsureNotSealed();
            if (_templateSource != null)
                throw new NotSupportedException($"{what} comes from the template for sheet '{Name}'. Change it in the template in Excel.");
            if (_headerWritten || (_bodyStarted && !_autoFit))
                throw new InvalidOperationException($"{what} must be defined before the first row is written. SheetStream streams the sheet forward-only and the OpenXML schema places this setting before the sheet data. Set it first, or enable AutoFitColumns() / use Cells[...] to defer the header.");
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
