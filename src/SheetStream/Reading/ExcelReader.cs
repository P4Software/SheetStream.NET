using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Xml;
using SheetStream.Internal;

namespace SheetStream.Reading
{
    /// <summary>An error value read from a cell, such as <c>#N/A</c> or <c>#DIV/0!</c>.</summary>
    public sealed class ExcelErrorValue
    {
        internal ExcelErrorValue(string code)
        {
            Code = code;
        }

        /// <summary>The error code as Excel shows it.</summary>
        public string Code { get; }

        /// <inheritdoc />
        public override string ToString() => Code;
    }

    /// <summary>
    /// One row read from a sheet. <see cref="Values"/> is zero-based like a DataRow (column A is index 0) and
    /// runs up to the last non-empty cell; missing cells are null.
    /// </summary>
    public sealed class ExcelRow
    {
        internal ExcelRow(int index, object?[] values)
        {
            Index = index;
            Values = values;
        }

        /// <summary>1-based row number in the sheet.</summary>
        public int Index { get; }

        /// <summary>
        /// Cell values: <see cref="string"/>, <see cref="double"/>, <see cref="bool"/>, <see cref="DateTime"/>
        /// (date-formatted numbers), <see cref="TimeSpan"/> (elapsed-time formats), <see cref="ExcelErrorValue"/>,
        /// or null. Formula cells yield their cached result.
        /// </summary>
        public IReadOnlyList<object?> Values { get; }

        /// <summary>Value at a zero-based column index, or null beyond the last cell.</summary>
        public object? this[int column] => column >= 0 && column < Values.Count ? Values[column] : null;

        /// <summary>Value as text (numbers in invariant culture), or null.</summary>
        public string? GetString(int column)
        {
            var v = this[column];
            return v switch
            {
                null => null,
                string s => s,
                double d => XmlTextUtil.ShortestRoundTrip(d),
                DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture),
                IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
                _ => v.ToString(),
            };
        }

        /// <summary>Value as a number; text is parsed with the invariant culture. Null when empty.</summary>
        public double? GetDouble(int column) => this[column] switch
        {
            null => null,
            double d => d,
            bool b => b ? 1 : 0,
            DateTime dt => ExcelDate.ToSerial(dt),
            TimeSpan ts => ts.TotalDays,
            string s when double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var p) => p,
            var other => throw new FormatException($"Row {Index}, column {CellReference.ColumnLetters(column + 1)}: '{other}' is not a number."),
        };

        /// <summary>Value as a date; serial numbers are converted. Null when empty.</summary>
        public DateTime? GetDateTime(int column) => this[column] switch
        {
            null => null,
            DateTime dt => dt,
            // FromOADate throws ArgumentException outside this range; a number that is not a date is a
            // conversion failure like any other, and is reported the same way.
            double d when d >= -657434 && d < 2958466 => ExcelDate.FromSerial(d),
            string s when DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var p) => p,
            var other => throw new FormatException($"Row {Index}, column {CellReference.ColumnLetters(column + 1)}: '{other}' is not a date."),
        };
    }

    /// <summary>
    /// Forward-only, low-memory reader for .xlsx files. Rows are parsed with <see cref="XmlReader"/> as they are
    /// enumerated, so a sheet of any size reads in constant memory apart from the shared string table.
    /// </summary>
    /// <example>
    /// <code>
    /// using var reader = ExcelReader.Open("import.xlsx");
    /// foreach (var row in reader.ReadRows("Products"))
    ///     Console.WriteLine($"{row.GetString(0)} = {row.GetDouble(2)}");
    /// var products = reader.Read&lt;Product&gt;("Products").ToList();
    /// </code>
    /// </example>
    public sealed partial class ExcelReader : IDisposable
    {
        private readonly ZipArchive _zip;
        private readonly Stream _stream;
        private readonly bool _leaveOpen;
        private readonly List<(string Name, string Part, SheetVisibility Visibility)> _sheets = new List<(string, string, SheetVisibility)>();
        private readonly List<string> _sharedStrings = new List<string>();
        private readonly List<CellKind> _styleKinds = new List<CellKind>();
        private readonly ReaderOptions _options;
        private bool _date1904;

        internal enum CellKind : byte { Number, Date, Duration }

        /// <summary>
        /// Deepest element nesting accepted in any part. SpreadsheetML nests about a dozen levels at most; the
        /// XML parser keeps a stack entry per open element, so a part made of millions of unclosed tags would
        /// otherwise grow that stack without bound.
        /// </summary>
        internal const int MaxXmlDepth = 256;

        /// <summary>Most merged ranges or hyperlinks returned for one sheet; a list beyond this is not a genuine workbook.</summary>
        internal const int MaxSheetAnnotations = 1_000_000;

        // Everything this reader keeps for its whole life (shared strings, style kinds, sheet list) is charged here.
        private readonly MemoryBudget _budget;

        private ExcelReader(Stream stream, bool leaveOpen, ReaderOptions? options)
        {
            _stream = stream;
            _leaveOpen = leaveOpen;
            _options = options ?? ReaderOptions.Default;
            _budget = new MemoryBudget(_options.MaxLoadedPartBytes);
            try
            {
                _zip = ZipDirectory.Open(stream);
            }
            catch (InvalidDataException ex)
            {
                if (!leaveOpen) stream.Dispose();
                throw new InvalidDataException("The file is not an .xlsx workbook (it is not a valid zip package). Old .xls files are not supported; save them as .xlsx first.", ex);
            }
            try
            {
                LoadWorkbook();
            }
            catch
            {
                // The caller never gets a reader to dispose, so what the constructor took over is released here.
                _zip.Dispose();
                if (!leaveOpen) stream.Dispose();
                throw;
            }
        }

        /// <summary>Opens a workbook file for reading. The file is shared for reading so Excel can keep it open.</summary>
        public static ExcelReader Open(string path) => Open(path, (ReaderOptions?)null);

        /// <summary>Opens a workbook file for reading with explicit safety limits; see <see cref="ReaderOptions"/>.</summary>
        public static ExcelReader Open(string path, ReaderOptions? options)
        {
            var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 16);
            return new ExcelReader(fs, false, options); // The constructor disposes the stream if it fails.
        }

        /// <summary>
        /// Opens a workbook from a stream. A non-seekable stream (an HTTP upload, say) is first copied, because
        /// the zip directory sits at the end of the file: into memory up to 64 MB, and into a temporary file
        /// that is deleted with the reader beyond that, so a large upload never sits in memory whole.
        /// </summary>
        public static ExcelReader Open(Stream stream, bool leaveOpen = false) => Open(stream, leaveOpen, null);

        /// <summary>Opens a workbook from a stream with explicit safety limits; see <see cref="ReaderOptions"/>.</summary>
        public static ExcelReader Open(Stream stream, bool leaveOpen, ReaderOptions? options)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (stream.CanSeek) return new ExcelReader(stream, leaveOpen, options);
            Stream copy;
            try
            {
                copy = Buffer(stream, InMemoryBufferLimit);
            }
            finally
            {
                // Released whether the copy worked or the source failed half-way: either way it is used up.
                if (!leaveOpen) stream.Dispose();
            }
            return new ExcelReader(copy, false, options); // On failure the constructor disposes the copy, which also deletes a spilled temp file.
        }

        /// <summary>
        /// How much of a non-seekable stream is kept in memory before the rest spills into a temporary file:
        /// 64 MB. Below the limit the whole upload sits in a <see cref="MemoryStream"/>; above it the bytes
        /// already buffered plus everything that follows go to a file opened with
        /// <see cref="FileOptions.DeleteOnClose"/>, so the file disappears when the reader is disposed.
        /// </summary>
        internal const long InMemoryBufferLimit = 64L * 1024 * 1024;

        /// <summary>
        /// Copies a non-seekable stream into a seekable one positioned at zero: a <see cref="MemoryStream"/>
        /// while the content fits in <paramref name="memoryLimit"/> bytes, otherwise a self-deleting temporary
        /// file. The synchronous twin of <c>BufferAsync</c>; internal so tests can use a small limit.
        /// </summary>
        internal static Stream Buffer(Stream source, long memoryLimit)
        {
            var memory = new MemoryStream();
            var buffer = new byte[81920]; // Below the large-object threshold.
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                if (memory.Length + read > memoryLimit)
                {
                    // GetRandomFileName + CreateNew instead of GetTempFileName: the latter caps at 65,535
                    // files per folder and creates the file twice.
                    var path = Path.Combine(Path.GetTempPath(), "sheetstream-" + Path.GetRandomFileName());
                    var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 1 << 16, FileOptions.DeleteOnClose);
                    try
                    {
                        memory.WriteTo(file);
                        file.Write(buffer, 0, read);
                        source.CopyTo(file);
                        file.Flush();
                        file.Position = 0;
                        return file;
                    }
                    catch
                    {
                        file.Dispose();
                        throw;
                    }
                }
                memory.Write(buffer, 0, read);
            }
            memory.Position = 0;
            return memory;
        }

        /// <summary>Sheet names in tab order (hidden sheets included).</summary>
        public IReadOnlyList<string> SheetNames
        {
            get
            {
                var names = new List<string>(_sheets.Count);
                foreach (var s in _sheets) names.Add(s.Name);
                return names;
            }
        }

        /// <summary>Visibility of a sheet, by name.</summary>
        public SheetVisibility GetVisibility(string sheetName) => _sheets[FindSheet(sheetName)].Visibility;

        /// <summary>Streams the rows of a sheet by name (case-insensitive).</summary>
        public IEnumerable<ExcelRow> ReadRows(string sheetName) => ReadRows(FindSheet(sheetName));

        /// <summary>Streams the rows of a sheet by zero-based position.</summary>
        public IEnumerable<ExcelRow> ReadRows(int sheetIndex) => ReadRowsCore(SheetEntry(sheetIndex));

        /// <summary>Resolves a sheet position to its zip entry, failing eagerly (not on the first MoveNext) for a bad index or a damaged package.</summary>
        private ZipArchiveEntry SheetEntry(int sheetIndex)
        {
            if (sheetIndex < 0 || sheetIndex >= _sheets.Count) throw new ArgumentOutOfRangeException(nameof(sheetIndex), sheetIndex, $"The workbook has {_sheets.Count} sheet(s).");
            return _zip.GetEntry(_sheets[sheetIndex].Part)
                ?? throw new InvalidDataException($"Sheet '{_sheets[sheetIndex].Name}' points at missing part '{_sheets[sheetIndex].Part}'.");
        }

        /// <summary>
        /// Reads a sheet into a <see cref="DataTable"/>. With <paramref name="firstRowIsHeader"/> the first row
        /// supplies column names (duplicates get a numeric suffix); otherwise columns are named A, B, C...
        /// Column types are object so mixed columns survive.
        /// </summary>
        /// <remarks>
        /// The table holds every row, so unlike <see cref="ReadRows(string)"/> its memory grows with the sheet.
        /// What it holds is counted against <see cref="ReaderOptions.MaxLoadedPartBytes"/> and a sheet that
        /// would take more is refused with an <see cref="InvalidDataException"/>: a table is rows times
        /// columns, so a few thousand rows that each have one cell in the last column would otherwise cost
        /// gigabytes. For a sheet too large for that, read it with <see cref="ReadRows(string)"/>.
        /// </remarks>
        public DataTable ReadDataTable(string sheetName, bool firstRowIsHeader = true)
        {
            var table = new DataTable(sheetName);
            var names = new ColumnNames(SharedStringIdentities());
            var held = new MemoryBudget(_options.MaxLoadedPartBytes, $"A DataTable of sheet '{sheetName}'");
            var first = true;
            foreach (var row in ReadRows(sheetName))
            {
                if (first && firstRowIsHeader)
                {
                    first = false;
                    AddHeaderColumns(table, row, names, held);
                    continue;
                }
                first = false;
                AddDataRow(table, row, names, held);
            }
            return table;
        }

        private HashSet<string>? _sharedIdentities;

        /// <summary>
        /// The shared string objects as a set compared by reference, built on first use. It lets
        /// <see cref="ReadDataTable(string, bool)"/> tell a cell that points at a shared string, which costs the
        /// table nothing more, from a cell with a text of its own.
        /// </summary>
        private HashSet<string> SharedStringIdentities()
        {
            if (_sharedIdentities != null) return _sharedIdentities;
            var set = new HashSet<string>(ByReference.Instance);
            foreach (var s in _sharedStrings) set.Add(s);
            _budget.Charge(40L * set.Count); // The set itself lives as long as the reader.
            return _sharedIdentities = set;
        }

        /// <summary>Compares strings by object identity. netstandard2.0 has no ReferenceEqualityComparer.</summary>
        private sealed class ByReference : IEqualityComparer<string>
        {
            public static readonly ByReference Instance = new ByReference();
            public bool Equals(string? x, string? y) => ReferenceEquals(x, y);
            public int GetHashCode(string obj) => System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
        }

        /// <summary>Turns the header row into table columns: blank headers get the column letter, duplicates a numeric suffix.</summary>
        private static void AddHeaderColumns(DataTable table, ExcelRow row, ColumnNames names, MemoryBudget held)
        {
            for (var i = 0; i < row.Values.Count; i++)
            {
                var name = row.GetString(i);
                table.Columns.Add(names.Unique(string.IsNullOrWhiteSpace(name) ? CellReference.ColumnLetters(i + 1) : name!), typeof(object));
            }
            held.Charge(TableColumnCost * (long)table.Columns.Count);
        }

        /// <summary>What a column of a <see cref="DataTable"/> is charged against the memory budget, before any rows.</summary>
        private const int TableColumnCost = 1024;

        /// <summary>What a row of a <see cref="DataTable"/> costs besides its cells: the row object and the table's own records of it.</summary>
        private const int TableRowCost = 200;

        /// <summary>
        /// What one cell of a <see cref="DataTable"/> costs, empty or not. A reference is eight bytes, and the
        /// table stores each column in an array it doubles as it grows, so on average half as much again sits
        /// unused. Measured, not assumed: a table of 3,000 rows by 16,384 empty cells takes about 1 GB.
        /// </summary>
        private const int TableSlotCost = 20;

        /// <summary>
        /// Hands out column names that are unique without regard to case. A name already taken gets the next
        /// free number appended, remembered per name, so a header row of ten thousand identical texts costs ten
        /// thousand lookups and not fifty million.
        /// </summary>
        private sealed class ColumnNames
        {
            private readonly HashSet<string> _used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, int> _next = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            private readonly HashSet<string> _shared;

            /// <param name="shared">The reader's shared strings, by identity: the same objects the cells hand out.</param>
            public ColumnNames(HashSet<string> shared)
            {
                _shared = shared;
            }

            /// <summary>True when <paramref name="text"/> is one of the reader's shared string objects, not a copy with the same characters.</summary>
            public bool IsShared(string text) => _shared.Contains(text);

            public string Unique(string name)
            {
                if (_used.Add(name)) return name;
                _next.TryGetValue(name, out var n);
                if (n < 2) n = 2;
                string candidate;
                do candidate = name + n++.ToString(CultureInfo.InvariantCulture);
                while (!_used.Add(candidate));
                _next[name] = n;
                return candidate;
            }
        }

        /// <summary>Appends a sheet row to the table, growing the column list when a row is wider than the header.</summary>
        private static void AddDataRow(DataTable table, ExcelRow row, ColumnNames names, MemoryBudget held)
        {
            var before = table.Columns.Count;
            // A column added here takes its letter as its name, which a header may already have used as a text.
            while (table.Columns.Count < row.Values.Count)
                table.Columns.Add(names.Unique(CellReference.ColumnLetters(table.Columns.Count + 1)), typeof(object));
            if (table.Columns.Count > before)
                held.Charge((table.Columns.Count - before) * (TableColumnCost + (long)TableSlotCost * table.Rows.Count)); // A new column has a slot for every row already there.

            var values = new object[table.Columns.Count];
            long text = 0;
            for (var i = 0; i < values.Length; i++)
            {
                var value = row[i];
                values[i] = value ?? DBNull.Value;
                // A text that came from the shared string table is one string however many cells show it,
                // and the reader has already paid for it. Charging it per cell refused ordinary sheets whose
                // text columns repeat a few values.
                if (value is string s) { if (!names.IsShared(s)) text += 2L * s.Length + 32; }
                else if (value != null) text += 24; // A boxed number, date or boolean.
            }
            held.Charge((long)TableSlotCost * values.Length + TableRowCost + text);
            table.Rows.Add(values);
        }

        /// <summary>
        /// Maps rows to objects. The first row is the header; each header is matched (case-insensitively) to a
        /// writable property or field by <see cref="SheetColumnAttribute"/> name, [DisplayName] or member name.
        /// Values are converted to the member type; a value that cannot be converted throws a
        /// <see cref="FormatException"/> naming the row and column. Unmatched columns are ignored.
        /// </summary>
        public IEnumerable<T> Read<[DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] T>(string sheetName) where T : new()
        {
            // A compile-time mapper (see ExcelReader.Mapping.cs) takes precedence over reflection.
            if (SheetMappers.TryGet<T>(out var mapper)) { foreach (var mapped in Read(sheetName, mapper)) yield return mapped; yield break; }
            Binding[]? bindings = null;
            foreach (var row in ReadRows(sheetName))
            {
                if (bindings == null)
                {
                    bindings = Bind(typeof(T), row);
                    continue;
                }
                if (MapRow(row, bindings, out T item)) yield return item;
            }
        }

        /// <summary>
        /// Fills a new <typeparamref name="T"/> from one data row through reflection bindings. Returns false for a
        /// blank row (no bound column has a value), which callers skip. Shared by the sync and async readers.
        /// </summary>
        private static bool MapRow<T>(ExcelRow row, Binding[] bindings, out T item) where T : new()
        {
            object boxed = new T()!;
            var any = false;
            foreach (var b in bindings)
            {
                var raw = row[b.Column];
                if (raw == null) continue;
                any = true;
                object? converted;
                try { converted = Convert(raw, b.Type); }
                catch (Exception ex) when (ex is FormatException || ex is InvalidCastException || ex is OverflowException || ex is ArgumentException)
                {
                    throw new FormatException($"Row {row.Index}, column {CellReference.ColumnLetters(b.Column + 1)} ('{b.Header}'): cannot convert '{raw}' to {b.Type.Name}.", ex);
                }
                b.Set(boxed, converted);
            }
            item = (T)boxed;
            return any;
        }

        /// <inheritdoc />
        public void Dispose()
        {
            _zip.Dispose();
            if (!_leaveOpen) _stream.Dispose();
        }

        private int FindSheet(string name)
        {
            for (var i = 0; i < _sheets.Count; i++)
                if (string.Equals(_sheets[i].Name, name, StringComparison.OrdinalIgnoreCase)) return i;
            throw new ArgumentException($"There is no sheet named '{name}'. Sheets: {string.Join(", ", SheetNames)}.", nameof(name));
        }

        // ------------------------------------------------------------------ package structure

        private void LoadWorkbook()
        {
            var workbookPart = "xl/workbook.xml";
            foreach (var rel in ReadRelationships("_rels/.rels", _budget))
                if (rel.Type.EndsWith("/officeDocument", StringComparison.Ordinal)) workbookPart = Resolve("", rel.Target);

            var folder = workbookPart.Contains("/") ? workbookPart.Substring(0, workbookPart.LastIndexOf('/') + 1) : "";
            var relsPart = folder + "_rels/" + workbookPart.Substring(folder.Length) + ".rels";
            var targets = new Dictionary<string, string>(StringComparer.Ordinal);
            string? sharedStringsPart = null, stylesPart = null;
            foreach (var rel in ReadRelationships(relsPart, _budget))
            {
                var target = Resolve(folder, rel.Target);
                targets[rel.Id] = target;
                if (rel.Type.EndsWith("/sharedStrings", StringComparison.Ordinal)) sharedStringsPart = target;
                else if (rel.Type.EndsWith("/styles", StringComparison.Ordinal)) stylesPart = target;
            }

            using (var xr = OpenXml(workbookPart) ?? throw new InvalidDataException("The package has no workbook part."))
            {
                while (Advance(xr))
                {
                    if (xr.NodeType != XmlNodeType.Element) continue;
                    // Only the workbook's own workbookPr, a direct child in the main namespace. Excel also writes
                    // an x15:workbookPr inside an extension list further down, which has no date1904 attribute;
                    // matching on the local name alone let that one switch the 1904 date system back off.
                    if (xr.LocalName == "workbookPr" && xr.Depth == 1 && xr.NamespaceURI == OpenXmlNamespaces.Main)
                    {
                        var d = xr.GetAttribute("date1904");
                        _date1904 = d == "1" || string.Equals(d, "true", StringComparison.OrdinalIgnoreCase);
                    }
                    else if (xr.LocalName == "sheet" && xr.NamespaceURI == OpenXmlNamespaces.Main)
                    {
                        var name = xr.GetAttribute("name") ?? "";
                        var rid = xr.GetAttribute("id", OpenXmlNamespaces.Relationships) ?? "";
                        var state = xr.GetAttribute("state");
                        var vis = state == "hidden" ? SheetVisibility.Hidden : state == "veryHidden" ? SheetVisibility.VeryHidden : SheetVisibility.Visible;
                        if (targets.TryGetValue(rid, out var part))
                        {
                            _budget.Charge(2L * (name.Length + part.Length) + 64);
                            _sheets.Add((name, part, vis));
                        }
                    }
                }
            }

            if (sharedStringsPart != null) LoadSharedStrings(sharedStringsPart);
            if (stylesPart != null) LoadStyles(stylesPart);
        }

        private void LoadSharedStrings(string part)
        {
            using (var xr = OpenXml(part))
            {
                if (xr == null) return;
                // Each item is bounded by the cell text limit, and the table as a whole twice over: by its
                // unpacked size, and by what its strings take in memory, which is charged as they are added.
                var text = new TextState(_options.MaxCellTextLength, long.MaxValue);
                while (!xr.EOF)
                {
                    // Depth 1: the table's own items, which are what cell indexes count.
                    if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "si" || xr.Depth != 1) { Advance(xr); continue; }
                    text.BeginCell();
                    var item = xr.IsEmptyElement ? "" : ReadRichText(xr, text); // Leaves the reader on </si>.
                    _budget.ChargeText(item);
                    _sharedStrings.Add(item);
                    Advance(xr);
                }
            }
        }

        /// <summary>
        /// Builds a per-style lookup of how numbers should surface: plain, date or elapsed time. Built-in ids
        /// 14-22 and 45-47 are dates and times; custom codes are inspected after removing quoted text,
        /// escapes and colour or condition brackets.
        /// </summary>
        private void LoadStyles(string part)
        {
            var custom = new Dictionary<int, string>();
            var kinds = new Dictionary<int, CellKind>();
            using (var xr = OpenXml(part))
            {
                if (xr == null) return;
                var inCellXfs = false;
                while (Advance(xr))
                {
                    if (xr.NodeType == XmlNodeType.Element)
                    {
                        if (xr.LocalName == "numFmt" && int.TryParse(xr.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
                        {
                            var code = xr.GetAttribute("formatCode") ?? "";
                            _budget.ChargeText(code);
                            _budget.Charge(32); // The dictionary entry.
                            custom[id] = code;
                        }
                        else if (xr.LocalName == "cellXfs") inCellXfs = !xr.IsEmptyElement;
                        else if (inCellXfs && xr.LocalName == "xf")
                        {
                            int.TryParse(xr.GetAttribute("numFmtId"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var fmt);
                            _budget.Charge(8); // One kind, with room for the list's spare capacity.
                            _styleKinds.Add(Classify(fmt, custom, kinds));
                        }
                    }
                    else if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "cellXfs")
                    {
                        inCellXfs = false;
                    }
                }
            }
        }

        /// <summary>
        /// <see cref="Classify(int, Dictionary{int, string})"/> with the answer remembered per format id.
        /// Inspecting a format code costs time in proportion to its length, and a workbook can point thousands
        /// of cell formats, or every cell of an edited sheet, at one very long code.
        /// </summary>
        internal static CellKind Classify(int id, Dictionary<int, string> custom, Dictionary<int, CellKind> known)
        {
            if (!known.TryGetValue(id, out var kind)) known[id] = kind = Classify(id, custom);
            return kind;
        }

        internal static CellKind Classify(int id, Dictionary<int, string> custom)
        {
            if (id == 46) return CellKind.Duration;
            if ((id >= 14 && id <= 22) || id == 45 || id == 47) return CellKind.Date;
            if (!custom.TryGetValue(id, out var code)) return CellKind.Number;

            var sb = new StringBuilder();
            var elapsed = false;
            for (var i = 0; i < code.Length; i++)
            {
                var c = code[i];
                if (c == '"') { i++; while (i < code.Length && code[i] != '"') i++; continue; }
                if (c == '\\' || c == '_' || c == '*') { i++; continue; }
                if (c == '[')
                {
                    var end = code.IndexOf(']', i);
                    if (end < 0) break;
                    var inner = code.Substring(i + 1, end - i - 1).ToLowerInvariant();
                    if (inner.Length > 0 && (inner[0] == 'h' || inner[0] == 'm' || inner[0] == 's') && inner.Trim('h', 'm', 's').Length == 0) elapsed = true;
                    i = end;
                    continue;
                }
                if (c == ';') break; // Only the positive section decides.
                sb.Append(char.ToLowerInvariant(c));
            }
            if (elapsed) return CellKind.Duration;
            var s = sb.ToString();
            if (s.Contains("general")) return CellKind.Number;
            foreach (var ch in s)
                if (ch == 'd' || ch == 'm' || ch == 'y' || ch == 'h' || ch == 's') return CellKind.Date;
            return CellKind.Number;
        }

        /// <summary>Reads a relationships part into a list, charging what the list holds to <paramref name="budget"/>.</summary>
        private IEnumerable<(string Id, string Type, string Target)> ReadRelationships(string part, MemoryBudget budget)
        {
            var list = new List<(string, string, string)>();
            using (var xr = OpenXml(part))
            {
                if (xr == null) return list;
                while (Advance(xr))
                {
                    if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "Relationship") continue;
                    var id = xr.GetAttribute("Id") ?? "";
                    var type = xr.GetAttribute("Type") ?? "";
                    var target = xr.GetAttribute("Target") ?? "";
                    budget.Charge(2L * (id.Length + type.Length + target.Length) + 128);
                    list.Add((id, type, target));
                }
            }
            return list;
        }

        /// <summary>Resolves a relationship target against the source part's folder ("worksheets/sheet1.xml", "/xl/x.xml", "../x.xml").</summary>
        private static string Resolve(string folder, string target) => TemplateSource.Resolve(folder, target);

        /// <summary>
        /// Opens a part that is loaded into memory as a whole (workbook, relationships, shared strings,
        /// styles), capped at <see cref="ReaderOptions.MaxLoadedPartBytes"/>. Returns null when the part does not exist.
        /// </summary>
        private XmlReader? OpenXml(string part)
        {
            var entry = _zip.GetEntry(part);
            if (entry == null) return null;
            return XmlReader.Create(Guard(entry, held: true), SheetSettings(async: false));
        }

        /// <summary>
        /// Decodes the unpacked bytes of a part under the limits of <see cref="GuardedPartReader"/>. A part
        /// that is <paramref name="held"/> in memory is capped by <see cref="ReaderOptions.MaxLoadedPartBytes"/>;
        /// a sheet, which is streamed, by <see cref="ReaderOptions.MaxStreamedPartBytes"/>.
        /// </summary>
        private TextReader Guard(ZipArchiveEntry entry, bool held)
            => new GuardedPartReader(entry.Open(), entry.FullName, held ? _options.MaxLoadedPartBytes : _options.MaxStreamedPartBytes, held, _options.MaxTokenChars, _options.MaxTagChars);

        // ------------------------------------------------------------------ rows and cells

        /// <summary>
        /// Settings for parsing a sheet part. DTDs are never resolved (untrusted input), and the reader owns the
        /// zip entry stream so disposing it closes both. <paramref name="async"/> enables the <c>*Async</c>
        /// methods used by <c>ReadRowsAsync</c>; the sync reader leaves it off because it costs a little per call.
        /// </summary>
        private static XmlReaderSettings SheetSettings(bool async)
            => new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                CloseInput = true,
                // Comments and processing instructions carry nothing a workbook reader needs; ignoring them
                // also means the parser skips their content instead of buffering it.
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                // Bounds the one thing the parser keeps for the whole part: its table of names.
                NameTable = new BoundedNameTable(),
                Async = async,
            };

        /// <summary>
        /// Moves to the next node and rejects absurd nesting. Every loop over untrusted XML goes through this
        /// (or its async twin) instead of calling <see cref="XmlReader.Read"/> directly.
        /// </summary>
        internal static bool Advance(XmlReader xr)
        {
            var more = xr.Read();
            if (xr.Depth > MaxXmlDepth) ThrowTooDeep();
            return more;
        }

        private static void ThrowTooDeep()
            => throw new InvalidDataException($"The workbook XML is nested more than {MaxXmlDepth} levels deep. The file is damaged or not a genuine workbook.");

        /// <summary>
        /// Skips the element the reader is on, with its whole subtree, and leaves the reader on the node after
        /// it. Same contract as <see cref="XmlReader.Skip"/>, but depth-checked.
        /// </summary>
        internal static void SkipElement(XmlReader xr)
        {
            if (!xr.IsEmptyElement)
            {
                var depth = xr.Depth;
                while (Advance(xr) && xr.Depth > depth) { }
            }
            Advance(xr); // Past the end tag, or past the empty element.
        }

        /// <summary>
        /// Scratch space and running totals for reading text out of untrusted XML. One instance serves a whole
        /// enumeration, so reading a cell allocates nothing but its final string.
        /// </summary>
        internal sealed class TextState
        {
            private readonly long _maxCell;
            private readonly long _maxRow;
            private long _cell;
            private long _row;

            public TextState(long maxCellChars, long maxRowChars)
            {
                _maxCell = maxCellChars;
                _maxRow = maxRowChars;
            }

            /// <summary>For parts of a workbook the application owns (templates), where no limit applies.</summary>
            public static TextState Unlimited() => new TextState(long.MaxValue, long.MaxValue);

            /// <summary>Chunk buffer for <see cref="XmlReader.ReadValueChunk"/>.</summary>
            public readonly char[] Buffer = new char[1024];

            /// <summary>Assembles a value that arrives in more than one chunk.</summary>
            public readonly StringBuilder Chunks = new StringBuilder();

            /// <summary>Assembles the runs of a rich-text item.</summary>
            public readonly StringBuilder Runs = new StringBuilder();

            public void BeginRow() => _row = 0;

            public void BeginCell() => _cell = 0;

            /// <summary>Accounts for characters just read and rejects a cell or row that has grown past its limit.</summary>
            public void Count(int chars)
            {
                _cell += chars;
                _row += chars;
                if (_cell > _maxCell)
                    throw new InvalidDataException($"A cell holds more than {_maxCell:N0} characters of text. The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxCellTextLength.");
                if (_row > _maxRow)
                    throw new InvalidDataException($"A row holds more than {_maxRow:N0} characters of text. The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxRowTextLength.");
            }
        }

        /// <summary>
        /// Reads the text of the element the reader is on (not empty) and leaves the reader on the node after
        /// its end tag, like <see cref="XmlReader.ReadElementContentAsString()"/>. Unlike that method the text
        /// is pulled in chunks and counted against the limits as it arrives, so an oversized value is rejected
        /// as soon as it passes its limit instead of after the whole of it has been materialised.
        /// </summary>
        internal static string ReadElementText(XmlReader xr, TextState state)
        {
            var depth = xr.Depth;
            string? single = null;
            StringBuilder? sb = null;
            Advance(xr);
            while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
            {
                switch (xr.NodeType)
                {
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        int n;
                        while ((n = xr.ReadValueChunk(state.Buffer, 0, state.Buffer.Length)) > 0)
                        {
                            state.Count(n);
                            if (single == null && sb == null) { single = new string(state.Buffer, 0, n); continue; }
                            if (sb == null) { sb = state.Chunks; sb.Clear(); sb.Append(single); }
                            sb.Append(state.Buffer, 0, n);
                        }
                        Advance(xr);
                        break;
                    case XmlNodeType.Element:
                        throw new InvalidDataException($"Element '{xr.LocalName}' is not allowed inside a cell value. The file is damaged.");
                    default:
                        Advance(xr);
                        break;
                }
            }
            Advance(xr);
            return sb != null ? sb.ToString() : single ?? "";
        }

        private IEnumerable<ExcelRow> ReadRowsCore(ZipArchiveEntry entry)
        {
            using (var xr = XmlReader.Create(Guard(entry, held: false), SheetSettings(async: false)))
            {
                var lastRow = 0;
                var values = new List<object?>();
                var text = new TextState(_options.MaxCellTextLength, _options.MaxRowTextLength);
                while (Advance(xr))
                {
                    if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "row") continue;
                    var rowIndex = ParseInt(xr.GetAttribute("r")) ?? lastRow + 1;
                    lastRow = rowIndex;
                    values.Clear();
                    text.BeginRow();
                    if (!xr.IsEmptyElement)
                    {
                        // Cursor discipline: every branch leaves the reader on the next unread node, so the
                        // loop never advances after a method that already did.
                        var depth = xr.Depth;
                        var nextColumn = 1;
                        Advance(xr);
                        while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
                        {
                            if (xr.NodeType == XmlNodeType.Element && xr.LocalName == "c")
                            {
                                var reference = xr.GetAttribute("r");
                                var column = reference != null ? ColumnOf(reference) : ImpliedColumn(nextColumn);
                                nextColumn = column + 1;
                                var value = ReadCell(xr, text); // Leaves the reader on </c> or on an empty <c/>.
                                Advance(xr);
                                if (value != null) PlaceCell(values, column, value);
                                continue;
                            }
                            Advance(xr);
                        }
                    }
                    yield return new ExcelRow(rowIndex, values.ToArray());
                }
            }
        }

        /// <summary>
        /// Column of a cell that carries no reference: the one after the previous cell. A row made of millions
        /// of such cells would otherwise grow without limit, so the same last column applies as for a reference.
        /// </summary>
        internal static int ImpliedColumn(int nextColumn)
        {
            if (nextColumn > CellReference.MaxColumns)
                throw new InvalidDataException("A row has more than 16,384 cells, which is past Excel's last column (XFD). The file is damaged or not a genuine workbook.");
            return nextColumn;
        }

        /// <summary>
        /// Stores a cell value at its 1-based column in the row buffer, padding skipped columns with null.
        /// ColumnOf and ImpliedColumn cap the index at 16384, so the padding is bounded even for hostile files.
        /// </summary>
        private static void PlaceCell(List<object?> values, int column, object value)
        {
            if (column - 1 < values.Count) { values[column - 1] = value; return; } // Out-of-order cell.
            while (values.Count < column - 1) values.Add(null);
            values.Add(value);
        }

        /// <summary>Reads one <c>c</c> element (positioned on its start) and returns its typed value.</summary>
        private object? ReadCell(XmlReader xr, TextState text)
        {
            var type = xr.GetAttribute("t");
            var styleIndex = ParseInt(xr.GetAttribute("s")) ?? 0;
            if (xr.IsEmptyElement) return null;

            string? v = null;
            string? inline = null;
            var depth = xr.Depth;
            text.BeginCell();
            Advance(xr);
            while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    if (xr.LocalName == "v")
                    {
                        if (xr.IsEmptyElement) { v = ""; Advance(xr); }
                        else v = ReadElementText(xr, text); // Advances past </v>.
                        continue;
                    }
                    if (xr.LocalName == "is")
                    {
                        if (xr.IsEmptyElement) inline = "";
                        else inline = ReadRichText(xr, text); // Leaves the reader on </is>.
                        Advance(xr);
                        continue;
                    }
                    SkipElement(xr); // <f> and anything else: skip the whole element.
                    continue;
                }
                Advance(xr);
            }

            return CellValue(type, styleIndex, v, inline);
        }

        /// <summary>
        /// Turns the raw parts of a <c>c</c> element (its <c>t</c> and <c>s</c> attributes, the <c>v</c> text and
        /// any inline string) into the typed value the row exposes. Shared by the sync and async parsers so the
        /// two can never disagree on how a cell reads.
        /// </summary>
        private object? CellValue(string? type, int styleIndex, string? v, string? inline)
        {
            switch (type)
            {
                case "s":
                    var idx = ParseInt(v);
                    return idx.HasValue && idx.Value >= 0 && idx.Value < _sharedStrings.Count ? _sharedStrings[idx.Value] : null;
                case "inlineStr":
                    return inline;
                case "str":
                    return v == null ? null : Decode(v);
                case "b":
                    return v == null ? (object?)null : v == "1" || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
                case "e":
                    return v == null ? null : new ExcelErrorValue(v);
                case "d":
                    return v != null && DateTime.TryParse(v, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var iso) ? iso : (object?)v;
                default:
                    if (v == null || !double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var number)) return v;
                    // double.TryParse accepts "NaN" and "Infinity", which no spreadsheet stores as a number;
                    // surface them as the text they are rather than as a value arithmetic cannot use.
                    if (double.IsNaN(number) || double.IsInfinity(number)) return v;
                    // A damaged file can carry s="-1"; treat any index outside the style table as a plain number.
                    var kind = styleIndex >= 0 && styleIndex < _styleKinds.Count ? _styleKinds[styleIndex] : CellKind.Number;
                    if (kind == CellKind.Date)
                    {
                        if (_date1904) number += 1462; // 1904 epoch: shift to the 1900 serial.
                        if (number >= -657435 && number < 2958466) return ExcelDate.FromSerial(number);
                    }
                    else if (kind == CellKind.Duration && Math.Abs(number) < MaxDurationDays)
                    {
                        return TimeSpan.FromDays(number);
                    }
                    return number; // Also a date or duration too large for its type: the raw number is the honest answer.
            }
        }

        /// <summary>Largest magnitude, in days, that fits a <see cref="TimeSpan"/> (its range is about 10,675,199 days).</summary>
        internal const double MaxDurationDays = 10_675_199;

        /// <summary>
        /// Concatenates the text of a string item (<c>si</c> or <c>is</c>): either one <c>t</c> or several runs.
        /// Phonetic guides (<c>rPh</c>) are skipped because they are not part of the cell text. Must be called
        /// on a non-empty start element and returns positioned on its end element.
        /// </summary>
        internal static string ReadRichText(XmlReader xr, TextState state)
        {
            // The common item is a single <t>; only build a string from runs when there is more than one.
            string? single = null;
            StringBuilder? sb = null;
            var depth = xr.Depth;
            Advance(xr);
            while (!xr.EOF && !(xr.NodeType == XmlNodeType.EndElement && xr.Depth == depth))
            {
                if (xr.NodeType == XmlNodeType.Element)
                {
                    if (xr.LocalName == "rPh") { SkipElement(xr); continue; }
                    if (xr.LocalName == "t")
                    {
                        if (xr.IsEmptyElement) { Advance(xr); continue; }
                        var run = ReadElementText(xr, state); // Advances past </t>.
                        if (single == null && sb == null) single = run;
                        else
                        {
                            if (sb == null) { sb = state.Runs; sb.Clear(); sb.Append(single); }
                            sb.Append(run);
                        }
                        continue;
                    }
                }
                Advance(xr);
            }
            return Decode(sb != null ? sb.ToString() : single ?? ""); // The reader is now on the container's end element.
        }

        /// <summary>Reverses Excel's <c>_xHHHH_</c> escaping of characters XML cannot carry.</summary>
        internal static string Decode(string s)
        {
            if (s.IndexOf("_x", StringComparison.Ordinal) < 0) return s;
            var sb = new StringBuilder(s.Length);
            for (var i = 0; i < s.Length; i++)
            {
                if (s[i] == '_' && i + 6 < s.Length && s[i + 1] == 'x' && s[i + 6] == '_'
                    && int.TryParse(s.Substring(i + 2, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var code))
                {
                    sb.Append((char)code);
                    i += 6;
                }
                else
                {
                    sb.Append(s[i]);
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Column number of an A1 reference. Rejects anything beyond Excel's 16,384 columns: an untrusted file
        /// could otherwise name column "ZZZZZZ" and make the row buffer grow to billions of entries.
        /// </summary>
        internal static int ColumnOf(string reference)
        {
            var column = 0;
            foreach (var ch in reference)
            {
                if (ch >= 'A' && ch <= 'Z') column = column * 26 + (ch - 'A' + 1);
                else if (ch >= 'a' && ch <= 'z') column = column * 26 + (ch - 'a' + 1);
                else break;
                if (column > CellReference.MaxColumns)
                    throw new InvalidDataException($"Cell reference '{reference}' is beyond Excel's last column (XFD). The file is damaged or not a genuine workbook.");
            }
            return column == 0 ? 1 : column;
        }

        private static int? ParseInt(string? s)
            => s != null && int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : (int?)null;

        // ------------------------------------------------------------------ object mapping

        private sealed class Binding
        {
            public Binding(int column, string header, Type type, Action<object, object?> set)
            {
                Column = column;
                Header = header;
                Type = type;
                Set = set;
            }

            public int Column { get; }
            public string Header { get; }
            public Type Type { get; }
            public Action<object, object?> Set { get; }
        }

        private static Binding[] Bind([DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties | DynamicallyAccessedMemberTypes.PublicFields)] Type type, ExcelRow header)
        {
            var byName = new Dictionary<string, MemberInfo>(StringComparer.OrdinalIgnoreCase);
            const BindingFlags flags = BindingFlags.Public | BindingFlags.Instance;
            // Properties then fields (not GetMembers) so the trimmer only has to keep those two member kinds.
            foreach (var m in type.GetProperties(flags).Concat<MemberInfo>(type.GetFields(flags)))
            {
                if (m is PropertyInfo p && (!p.CanWrite || p.GetIndexParameters().Length > 0)) continue;
                if (!(m is PropertyInfo) && !(m is FieldInfo f && !f.IsInitOnly)) continue;
                var attr = m.GetCustomAttribute<SheetColumnAttribute>();
                if (attr?.Ignore == true) continue;
                var display = m.GetCustomAttribute<System.ComponentModel.DisplayNameAttribute>()?.DisplayName;
                if (attr?.Name != null) byName[attr.Name] = m;
                if (display != null && !byName.ContainsKey(display)) byName[display] = m;
                if (!byName.ContainsKey(m.Name)) byName[m.Name] = m;
            }

            var bindings = new List<Binding>();
            for (var i = 0; i < header.Values.Count; i++)
            {
                var text = header.GetString(i)?.Trim();
                if (string.IsNullOrEmpty(text) || !byName.TryGetValue(text!, out var member)) continue;
                if (member is PropertyInfo prop)
                    bindings.Add(new Binding(i, text!, prop.PropertyType, (o, v) => prop.SetValue(o, v)));
                else if (member is FieldInfo field)
                    bindings.Add(new Binding(i, text!, field.FieldType, (o, v) => field.SetValue(o, v)));
            }
            return bindings.ToArray();
        }

        /// <summary>Converts a cell value to a member type, handling nullables, enums, dates and Guids.</summary>
        private static object? Convert(object value, Type target)
        {
            var type = Nullable.GetUnderlyingType(target) ?? target;
            if (type.IsInstanceOfType(value)) return value;
            if (type == typeof(string)) return value is double d ? XmlTextUtil.ShortestRoundTrip(d) : value is IFormattable f ? f.ToString(null, CultureInfo.InvariantCulture) : value.ToString();
            if (value is string s && s.Trim().Length == 0) return null;
            if (type == typeof(DateTime))
                return value switch
                {
                    double num => ExcelDate.FromSerial(num),
                    string str => DateTime.Parse(str, CultureInfo.InvariantCulture),
                    _ => throw new InvalidCastException(),
                };
            if (type == typeof(DateTimeOffset)) return new DateTimeOffset((DateTime)Convert(value, typeof(DateTime))!);
            if (type == typeof(TimeSpan))
                return value switch
                {
                    double num => TimeSpan.FromDays(num),
                    string str => TimeSpan.Parse(str, CultureInfo.InvariantCulture),
                    DateTime dt => dt.TimeOfDay,
                    _ => throw new InvalidCastException(),
                };
            if (type == typeof(Guid)) return Guid.Parse(value.ToString()!);
            if (type.IsEnum)
                return value is string es ? Enum.Parse(type, es.Trim(), ignoreCase: true) : Enum.ToObject(type, System.Convert.ToInt64(value, CultureInfo.InvariantCulture));
            if (type == typeof(bool) && value is string bs)
            {
                var t = bs.Trim();
                if (t == "1" || t.Equals("yes", StringComparison.OrdinalIgnoreCase) || t.Equals("y", StringComparison.OrdinalIgnoreCase)) return true;
                if (t == "0" || t.Equals("no", StringComparison.OrdinalIgnoreCase) || t.Equals("n", StringComparison.OrdinalIgnoreCase)) return false;
                return bool.Parse(t);
            }
            if (value is DateTime dtv && (type == typeof(double) || type == typeof(decimal))) value = ExcelDate.ToSerial(dtv);
            return System.Convert.ChangeType(value, type, CultureInfo.InvariantCulture);
        }
    }
}
