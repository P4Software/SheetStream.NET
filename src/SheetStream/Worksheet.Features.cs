using System;
using System.Collections.Generic;
using System.IO;
using SheetStream.Internal;

namespace SheetStream
{
    public sealed partial class Worksheet
    {
        private readonly List<HyperlinkEntry> _hyperlinks = new List<HyperlinkEntry>();
        private readonly List<DataValidation> _validations = new List<DataValidation>();
        private readonly List<ConditionalFormatting> _conditionalFormats = new List<ConditionalFormatting>();
        private readonly List<ExcelTable> _tables = new List<ExcelTable>();
        private readonly List<ExcelPicture> _pictures = new List<ExcelPicture>();
        private readonly List<ExcelChart> _charts = new List<ExcelChart>();
        private readonly List<CellComment> _comments = new List<CellComment>();
        private readonly List<PivotTable> _pivotTables = new List<PivotTable>();
        private SheetProtection? _protection;

        // AutoFit state: widest content seen per column, in Calibri-11 character units.
        private bool _autoFit;
        private double _autoFitMin;
        private double _autoFitMax;
        private double[] _widths = new double[0];

        // Widths of cells in the first rows, kept so that rows which later turn out to be table or filter
        // headers (LoadFromCollection adds its table after writing) can be padded for the drop-down button.
        private const int EarlyRowLimit = 20;
        private readonly Dictionary<long, double> _earlyWidths = new Dictionary<long, double>();

        /// <summary>Extra characters Excel reserves for a filter drop-down button in a header cell.</summary>
        private const double FilterButtonWidth = 2.5;

        /// <summary>A hyperlink attached to a cell.</summary>
        internal readonly struct HyperlinkEntry
        {
            public HyperlinkEntry(int row, int column, string? url, string? location, string? tooltip)
            {
                Row = row; Column = column; Url = url; Location = location; Tooltip = tooltip;
            }

            public int Row { get; }
            public int Column { get; }
            public string? Url { get; }
            public string? Location { get; }
            public string? Tooltip { get; }
        }

        /// <summary>
        /// Sizes columns to their content when the sheet is saved. Call before the first row. Widths are
        /// estimated from character counts (no font metrics are available without a graphics library), scaled
        /// for bold and larger fonts, and clamped to [<paramref name="minWidth"/>, <paramref name="maxWidth"/>].
        /// Columns given an explicit width keep it.
        /// </summary>
        /// <remarks>
        /// Column widths sit before the data in the file, so this mode writes the rows to a temporary file and
        /// copies them in at save time. Memory stays flat; disk use roughly equals the uncompressed sheet size.
        /// </remarks>
        public Worksheet AutoFitColumns(double minWidth = 6, double maxWidth = 80)
        {
            if (minWidth < 0 || maxWidth > 255 || minWidth > maxWidth) throw new ArgumentOutOfRangeException(nameof(maxWidth), "Widths must satisfy 0 <= minWidth <= maxWidth <= 255.");
            EnsureNotSealed();
            if (_templateSource != null) throw new NotSupportedException($"Sheet '{Name}' fills a template, so its column widths come from the template.");
            if (_bodyStarted) throw new InvalidOperationException("AutoFitColumns must be called before the first row is written, so the rows can be measured.");
            _autoFit = true;
            _autoFitMin = minWidth;
            _autoFitMax = maxWidth;
            return this;
        }

        /// <summary>True when RowWriter must report cell text for this row (AutoFit or a pending table header).</summary>
        internal bool ShouldObserveRow(int row)
        {
            if (_autoFit) return true;
            foreach (var t in _tables)
                if (t.Range.FromRow == row && t.NeedsHeaderCapture) return true;
            return false;
        }

        /// <summary>Receives the displayed text of a cell written on an observed row.</summary>
        internal void Observe(int row, int column, string text, int styleId)
        {
            if (_autoFit)
            {
                if (column >= _widths.Length)
                {
                    var grown = new double[Math.Min(CellReference.MaxColumns + 1, Math.Max(column + 1, _widths.Length * 2 + 8))];
                    Array.Copy(_widths, grown, _widths.Length);
                    _widths = grown;
                }
                var chars = LongestLine(text) * Workbook.Styles.WidthFactor(styleId);
                if (chars > _widths[column]) _widths[column] = chars;
                if (row <= EarlyRowLimit) _earlyWidths[(long)row << 16 | (uint)column] = chars;
            }
            foreach (var t in _tables)
                if (t.Range.FromRow == row) t.CaptureHeader(column, text);
        }

        private static int LongestLine(string text)
        {
            var longest = 0;
            var current = 0;
            foreach (var c in text)
            {
                if (c == '\n') { if (current > longest) longest = current; current = 0; }
                else if (c != '\r') current++;
            }
            return current > longest ? current : longest;
        }

        /// <summary>Turns measured character counts into column widths, leaving explicit widths alone.</summary>
        private void ApplyAutoFitWidths()
        {
            // Header cells of tables and AutoFilters need room for the drop-down button as well as the text.
            var headers = new List<(int Row, int From, int To)>();
            foreach (var t in _tables) if (t.ShowFilterButtons) headers.Add((t.Range.FromRow, t.Range.FromColumn, t.Range.ToColumn));
            if (_autoFilterRelative != null)
            {
                var f = RangeReference.Parse(_autoFilterRelative);
                headers.Add((f.FromRow, f.FromColumn, f.ToColumn));
            }
            foreach (var (row, from, to) in headers)
            {
                for (var c = from; c <= to && c < _widths.Length; c++)
                {
                    if (_earlyWidths.TryGetValue((long)row << 16 | (uint)c, out var w) && w + FilterButtonWidth > _widths[c])
                        _widths[c] = w + FilterButtonWidth;
                }
            }

            for (var c = 1; c < _widths.Length; c++)
            {
                if (_widths[c] <= 0) continue;
                if (_columns.TryGetValue(c, out var existing) && existing.HasExplicitWidth) continue;
                // About one character of padding plus a little for the filter button area, as Excel's own AutoFit leaves.
                var width = Math.Min(_autoFitMax, Math.Max(_autoFitMin, _widths[c] + 2));
                GetOrCreateColumn(c).SetAutoWidth(width);
            }
        }

        /// <summary>Records a hyperlink on a cell. Called by <see cref="RowWriter"/> and <see cref="AddHyperlink"/>.</summary>
        internal void RegisterHyperlink(int row, int column, string? url, string? location, string? tooltip)
        {
            EnsureNotSealed();
            if (url != null && !Uri.IsWellFormedUriString(url, UriKind.RelativeOrAbsolute))
                throw new ArgumentException($"'{url}' is not a valid URL. Encode spaces as %20.", nameof(url));
            if (_hyperlinks.Count >= 65530)
                throw new InvalidOperationException("Excel supports at most 65,530 hyperlinks per worksheet.");
            _hyperlinks.Add(new HyperlinkEntry(row, column, url, location, tooltip));
        }

        /// <summary>
        /// Attaches a web, mail or file hyperlink to a cell that is (or will be) written separately. To write a
        /// link cell in one call use <see cref="RowWriter.WriteHyperlink"/>.
        /// </summary>
        public Worksheet AddHyperlink(int row, int column, string url, string? tooltip = null)
        {
            if (string.IsNullOrWhiteSpace(url)) throw new ArgumentException("URL cannot be empty.", nameof(url));
            ValidateRow(row, nameof(row));
            CellReference.ColumnLetters(column);
            RegisterHyperlink(row, column, url, null, tooltip);
            return this;
        }

        /// <summary>
        /// Attaches a link to a place in this workbook, e.g. <c>"'Summary'!A1"</c> or a defined name.
        /// </summary>
        public Worksheet AddInternalLink(int row, int column, string location, string? tooltip = null)
        {
            if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException("Location cannot be empty.", nameof(location));
            ValidateRow(row, nameof(row));
            CellReference.ColumnLetters(column);
            RegisterHyperlink(row, column, null, location.TrimStart('#'), tooltip);
            return this;
        }

        /// <summary>
        /// Adds a data validation (drop-down list, number / date / length bounds, or custom formula) over an
        /// A1 range such as <c>"C2:C500"</c>. Pick the rule on the returned object.
        /// </summary>
        public DataValidation AddDataValidation(string range)
        {
            EnsureNotSealed();
            return new DataValidation(RangeReference.Parse(range), v => { EnsureNotSealed(); _validations.Add(v); });
        }

        /// <summary>Adds a data validation over a rectangle given by row and column numbers.</summary>
        public DataValidation AddDataValidation(int fromRow, int fromColumn, int toRow, int toColumn)
            => AddDataValidation(new RangeReference(fromRow, fromColumn, toRow, toColumn).ToString());

        /// <summary>Adds conditional formatting rules over an A1 range such as <c>"C2:C101"</c>.</summary>
        public ConditionalFormatting AddConditionalFormatting(string range)
        {
            EnsureNotSealed();
            var cf = new ConditionalFormatting(this, RangeReference.Parse(range));
            _conditionalFormats.Add(cf);
            return cf;
        }

        /// <summary>Adds conditional formatting over a rectangle given by row and column numbers.</summary>
        public ConditionalFormatting AddConditionalFormatting(int fromRow, int fromColumn, int toRow, int toColumn)
            => AddConditionalFormatting(new RangeReference(fromRow, fromColumn, toRow, toColumn).ToString());

        /// <summary>
        /// Protects the sheet so locked cells cannot be edited. Cells whose style sets
        /// <c>Protection.Locked = false</c> stay editable. The password is stored as a salted SHA-512 hash.
        /// Sheet protection deters accidental edits; it is not encryption.
        /// </summary>
        public Worksheet Protect(string? password = null, Action<SheetProtectionOptions>? configure = null)
        {
            EnsureNotSealed();
            var options = new SheetProtectionOptions();
            configure?.Invoke(options);
            _protection = new SheetProtection(password, options);
            return this;
        }

        /// <summary>
        /// Turns an A1 range (header row included) into an Excel table. Add it before writing the header row so
        /// the header texts can be captured, or pass <paramref name="columnNames"/>.
        /// </summary>
        /// <example>
        /// <code>
        /// var table = sheet.AddTable("Sales", "A1:C101");
        /// table.StyleName = "TableStyleMedium9";
        /// sheet.Row(1).WriteCell("ID").WriteCell("Product").WriteCell("Revenue");
        /// </code>
        /// </example>
        public ExcelTable AddTable(string name, string range, IReadOnlyList<string>? columnNames = null)
        {
            EnsureNotSealed();
            var r = RangeReference.Parse(range);
            if (r.RowCount < 2) throw new ArgumentException($"Table range {range} needs a header row and at least one data row.", nameof(range));
            foreach (var t in _tables)
                if (t.Range.Overlaps(r)) throw new ArgumentException($"Table '{name}' overlaps table '{t.Name}' ({t.Range}).", nameof(range));
            if (columnNames == null && (_headerCaptureImpossible(r.FromRow)))
                throw new InvalidOperationException($"Row {r.FromRow} has already been written, so the header texts of table '{name}' can no longer be captured. Add the table before writing its header row, or pass columnNames.");
            Workbook.ReserveTableName(name);
            var table = new ExcelTable(this, ++Workbook.TableCount, name, r, columnNames);
            _tables.Add(table);
            _row.RefreshObservation(); // The header row may already be open with no cells yet.
            return table;
        }

        /// <summary>Adds a table over a rectangle given by row and column numbers.</summary>
        public ExcelTable AddTable(string name, int fromRow, int fromColumn, int toRow, int toColumn, IReadOnlyList<string>? columnNames = null)
            => AddTable(name, new RangeReference(fromRow, fromColumn, toRow, toColumn).ToString(), columnNames);

        private bool _headerCaptureImpossible(int headerRow)
            => _buffer == null && (headerRow < CurrentRow || (headerRow == CurrentRow && _row.IsOpen && _row.NextColumn > 1));

        /// <summary>
        /// Places a PNG, JPEG or GIF image with its top-left corner in the given cell. Without a size the image
        /// keeps its pixel size; with only one dimension the other follows the aspect ratio.
        /// </summary>
        public ExcelPicture AddPicture(byte[] image, int row, int column, int? widthPixels = null, int? heightPixels = null, string? name = null)
        {
            if (image == null) throw new ArgumentNullException(nameof(image));
            EnsureNotSealed();
            ValidateRow(row, nameof(row));
            CellReference.ColumnLetters(column);
            var (ext, w, h) = ExcelPicture.Inspect(image);
            if (w <= 0 || h <= 0) throw new ArgumentException("The image reports a zero size.", nameof(image));
            int width, height;
            if (widthPixels.HasValue && heightPixels.HasValue) { width = widthPixels.Value; height = heightPixels.Value; }
            else if (widthPixels.HasValue) { width = widthPixels.Value; height = (int)Math.Round((double)h * width / w); }
            else if (heightPixels.HasValue) { height = heightPixels.Value; width = (int)Math.Round((double)w * height / h); }
            else { width = w; height = h; }
            if (width <= 0 || height <= 0) throw new ArgumentOutOfRangeException(nameof(widthPixels), "Picture size must be positive.");
            var pic = new ExcelPicture(image, ext, width, height, row, column, name ?? "Picture " + (_pictures.Count + 1));
            _pictures.Add(pic);
            return pic;
        }

        /// <summary>Places an image read from a file. See <see cref="AddPicture(byte[], int, int, int?, int?, string?)"/>.</summary>
        public ExcelPicture AddPicture(string path, int row, int column, int? widthPixels = null, int? heightPixels = null)
            => AddPicture(File.ReadAllBytes(path), row, column, widthPixels, heightPixels, Path.GetFileNameWithoutExtension(path));

        /// <summary>
        /// Adds a chart that fills the cells of <paramref name="placement"/> (e.g. <c>"E2:L20"</c>). Add series
        /// that reference data on any sheet; Excel draws the chart from the cells when the file opens.
        /// </summary>
        public ExcelChart AddChart(ChartType type, string placement)
        {
            EnsureNotSealed();
            var chart = new ExcelChart(type, RangeReference.Parse(placement), Name, ++Workbook.ChartCount);
            _charts.Add(chart);
            return chart;
        }

        /// <summary>
        /// Adds a pivot table to this sheet with its top-left cell at (<paramref name="row"/>, <paramref name="column"/>),
        /// summarising <paramref name="sourceRange"/> (header row included) on sheet <paramref name="sourceSheetName"/>.
        /// Excel computes the pivot when the file opens; see <see cref="PivotTable"/> for the design and its limits.
        /// </summary>
        /// <param name="name">Pivot table name, unique within the workbook (e.g. "SalesByRegion").</param>
        /// <param name="sourceSheetName">Sheet holding the data. It must already exist in this workbook (it may be this sheet).</param>
        /// <param name="sourceRange">A1 range of the data including its header row, e.g. "A1:F500".</param>
        /// <param name="row">1-based row of the pivot's top-left cell on this sheet.</param>
        /// <param name="column">1-based column of the pivot's top-left cell on this sheet.</param>
        /// <param name="headers">
        /// The header texts of the source columns, in order. The source sheet has already been streamed out, so
        /// SheetStream cannot read them back; <see cref="LoadResult.Headers"/> returns them for data written with
        /// a <c>LoadFrom*</c> call, and the <see cref="AddPivotTable(string, ExcelTable, int, int)"/> overload takes
        /// them from the table.
        /// </param>
        /// <example>
        /// <code>
        /// var loaded = data.LoadFromCollection(orders);
        /// var pivot = summary.AddPivotTable("ByRegion", "Data", loaded.Range, 1, 1, loaded.Headers);
        /// pivot.AddRowField("Region").AddDataField("Amount");
        /// </code>
        /// </example>
        public PivotTable AddPivotTable(string name, string sourceSheetName, string sourceRange, int row, int column, IReadOnlyList<string> headers)
        {
            if (headers == null) throw new ArgumentNullException(nameof(headers));
            if (string.IsNullOrWhiteSpace(sourceSheetName)) throw new ArgumentException("Source sheet name cannot be empty.", nameof(sourceSheetName));
            var range = RangeReference.Parse(sourceRange);
            if (range.RowCount < 2) throw new ArgumentException($"Pivot source {sourceRange} needs a header row and at least one data row.", nameof(sourceRange));
            if (headers.Count != range.ColumnCount)
                throw new ArgumentException($"Pivot source {sourceRange} spans {range.ColumnCount} columns but {headers.Count} headers were given.", nameof(headers));
            var source = FindSheet(sourceSheetName) ?? throw new ArgumentException($"This workbook has no sheet named '{sourceSheetName}'. Write the source sheet before adding the pivot table.", nameof(sourceSheetName));
            return AddPivotTableCore(name, source.Name, range, null, headers, row, column);
        }

        /// <summary>
        /// Adds a pivot table whose source is an Excel table written earlier (or on this sheet). The field names
        /// come from the table's columns, and Excel re-reads the whole table on every refresh, so the pivot follows
        /// the table as it grows.
        /// </summary>
        /// <param name="name">Pivot table name, unique within the workbook.</param>
        /// <param name="source">The table to summarise, from <see cref="AddTable(string, string, IReadOnlyList{string})"/> or <see cref="LoadOptions.TableName"/>.</param>
        /// <param name="row">1-based row of the pivot's top-left cell on this sheet.</param>
        /// <param name="column">1-based column of the pivot's top-left cell on this sheet.</param>
        public PivotTable AddPivotTable(string name, ExcelTable source, int row, int column)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (!ReferenceEquals(source.Sheet.Workbook, Workbook))
                throw new ArgumentException($"Table '{source.Name}' belongs to another workbook.", nameof(source));
            var headers = source.ColumnNames
                ?? throw new InvalidOperationException($"The header texts of table '{source.Name}' are not known yet. Write its header row (or pass column names to AddTable) before adding the pivot table.");
            return AddPivotTableCore(name, source.Sheet.Name, source.Range, source.Name, headers, row, column);
        }

        private PivotTable AddPivotTableCore(string name, string sourceSheet, RangeReference sourceRange, string? sourceTable, IReadOnlyList<string> headers, int row, int column)
        {
            EnsureNotSealed();
            if (_templateSource != null)
                throw new NotSupportedException($"Sheet '{Name}' fills a template or an edited workbook; pivot tables can only be added to sheets written from scratch.");
            ValidateRow(row, nameof(row));
            CellReference.ColumnLetters(column);
            // Construct first: the constructor validates the headers, and a failed call must not leave the
            // name reserved or a part number consumed.
            var pivot = new PivotTable(this, Workbook.PivotTableCount + 1, name, sourceSheet, sourceRange, sourceTable, headers, row, column);
            Workbook.ReservePivotTableName(name);
            Workbook.PivotTableCount++;
            _pivotTables.Add(pivot);
            return pivot;
        }

        /// <summary>Finds a worksheet of this workbook by name (case-insensitive, like Excel).</summary>
        private Worksheet? FindSheet(string name)
        {
            foreach (var s in Workbook.Worksheets)
                if (string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)) return s;
            return null;
        }

        /// <summary>
        /// Attaches a comment (Excel note) to a cell. Notes appear when the mouse hovers over the cell's red
        /// corner mark.
        /// </summary>
        public Worksheet AddComment(int row, int column, string text, string author = "SheetStream")
        {
            if (text == null) throw new ArgumentNullException(nameof(text));
            if (string.IsNullOrWhiteSpace(author)) throw new ArgumentException("Author cannot be empty.", nameof(author));
            EnsureNotSealed();
            ValidateRow(row, nameof(row));
            CellReference.ColumnLetters(column);
            if (_comments.Count >= 1023) throw new InvalidOperationException("SheetStream supports up to 1,023 comments per worksheet.");
            _comments.Add(new CellComment(row, column, text, author));
            return this;
        }
    }
}
