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

        // ---- Where the bytes go -------------------------------------------------------------------------------------
        //
        // A sheet's rows are UTF-8 composed by hand, in a byte buffer, because the XmlWriter's per-element and
        // per-attribute work (name checks, state machine, virtual calls, re-encoding characters) costs more than the
        // text of a cell. There are two targets for that buffer:
        //
        //  * Direct (new sheets): the buffer IS the write buffer of the zip entry (a PartStream), so a cell is formatted
        //    once, in place, and reaches the compressor in large blocks. The <row> element is written by hand too.
        //    The XmlWriter has finished the sheet's header and left the <sheetData> start tag open; it is closed
        //    just before the first row, and from then on the XmlWriter writes nothing until the rows are done.
        //  * XML (sheets filled from a template, or edited): rows are interleaved with elements the XmlWriter copies
        //    from the template, so the bytes are decoded back to text and handed to XmlWriter.WriteRaw, and the
        //    row element, with the attributes of the row it replaces, is written by the XmlWriter.
        //
        // Both targets are written by the same cell code. Escaping is done by hand in exactly the way XmlWriter
        // escapes text (see AppendText), and a test compares the result with a real XmlWriter, character for character.
        private PartStream? _stream;
        private XmlWriter? _w;
        private XmlWriter? _closeBeforeFirstRow;
        private byte[] _buf = Array.Empty<byte>();
        private int _n;

        // XML target only: the staging buffer is its own array (16 KB, dropped when the sheet is sealed) and its
        // contents are decoded into _chars. The decoder keeps state, so a flush may fall inside a character.
        internal const int XmlTargetBufferLength = 1 << 14;
        private char[]? _chars;
        private Decoder? _decoder;

        // A UTF-8 character is at most four bytes, so a flush that ended inside one leaves at most three behind.
        private const int MaxCarriedOverBytes = 3;

        // Room every composed cell needs on top of its text. The longest fixed part is an inline string with an
        // address, a style and xml:space (about 80 bytes) and a number is at most 40, so 192 is generous.
        // Text is the only variable part and is written in pieces that flush as needed.
        private const int CellReserve = 192;

        // The row number as text, formatted once per row for the r attribute of any cell that needs one.
        private readonly byte[] _rowDigits = new byte[Utf8Numbers.MaxLength];
        private int _rowDigitsLength;

        // True from the start of a row until its first cell: such a row is written as <row r="3" /> like XmlWriter did.
        private bool _rowEmpty;

        // Column of the last <c> element written in the open row, 0 before the first. A cell is consecutive,
        // and may leave out its r attribute, exactly when NextColumn is this plus one.
        private int _lastWrittenColumn;

        // PackageOptions.WriteCellReferences and the string storage mode, read once per row so each cell pays one test
        // and no chain of property reads.
        private bool _alwaysWriteReference;
        private StringStorage _stringStorage;
        private SharedStringTable? _sharedStrings;

        // StringStorage.Auto only: what is known about each column's text, indexed by column number. It is created
        // when the first text cell of the sheet is written, grows to the highest column that holds text (never to
        // more than 16,384 + 1 entries) and is released with the sheet. Its size depends on the number of columns,
        // never on the number of rows.
        private StringColumnState[] _columnStates = Array.Empty<StringColumnState>();

        // True when the sheet wants each cell's display text (AutoFit, or a table header being captured).
        // Computed once per row so the common path pays a single bool test per cell.
        private bool _observe;

        // True when the sheet caches formula results and therefore wants every value recorded. Read once per
        // row, so with the feature off (the default) each cell pays one false branch and nothing else.
        private bool _cache;

        private CellStyle? _rowStyle;
        private int _lastStyleId;

        // Row style combined with a date / date-time / duration format, derived once per row on first use.
        private CellStyle? _rowDateStyle;
        private CellStyle? _rowDateTimeStyle;
        private CellStyle? _rowDurationStyle;

        // Renders rich text runs, the one kind of cell that still needs an XmlWriter. See WriteRichText.
        private XmlWriter? _fragment;
        private MemoryStream? _fragmentBytes;

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

        /// <summary>
        /// Connects the writer to a stream whose buffer it composes rows in directly (new sheets). When
        /// <paramref name="partToClose"/> is given, it is the part's XML writer, which is left with an unfinished
        /// <c>&lt;sheetData&gt;</c> start tag that the first row has to close; it is flushed so everything it wrote is
        /// in the stream before the first byte of the first row.
        /// </summary>
        internal void AttachDirect(PartStream stream, XmlWriter? partToClose)
        {
            _stream = stream;
            _w = null;
            _closeBeforeFirstRow = partToClose;
        }

        /// <summary>Connects the writer to an XmlWriter that also carries other content around the rows (templates, edits).</summary>
        internal void AttachXml(XmlWriter writer)
        {
            _stream = null;
            _w = writer;
            _closeBeforeFirstRow = null;
        }

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
            // A failed write to the output is final (see PartStream). Said first, before any state changes, so a
            // caller that swallowed the failure keeps being told what went wrong instead of meeting whatever the
            // half-flushed buffer happens to hold.
            _stream?.ThrowIfFaulted();
            RowIndex = rowIndex;
            NextColumn = 1;
            _lastWrittenColumn = 0;
            var workbook = _sheet.Workbook;
            _alwaysWriteReference = workbook.Package.Options.WriteCellReferences;
            _stringStorage = workbook.StringStorage;
            _sharedStrings = workbook.SharedStrings;
            var style = options?.Style;
            _rowStyle = style != null && !style.IsDefault ? style : null;
            _rowDateStyle = _rowDateTimeStyle = _rowDurationStyle = null;
            _observe = _sheet.ShouldObserveRow(rowIndex);
            _cache = _sheet.IsCachingFormulaResults;
            _rowDigitsLength = Utf8Numbers.Format(rowIndex, _rowDigits, 0);

            // Checked before the row is marked open, so a bad height leaves nothing half written.
            if (options != null)
            {
                var h = options.Height;
                if (h.HasValue && (h.Value < 0 || h.Value > 409)) throw new ArgumentOutOfRangeException(nameof(options), h.Value, "Row height must be between 0 and 409 points.");
            }

            IsOpen = true;
            if (_stream != null) BeginDirect(options);
            else BeginXml(options, rawAttributes);
        }

        /// <summary>The row's own element, written by hand: <c>&lt;row r="3"&gt;</c> and the optional attributes.</summary>
        private void BeginDirect(RowOptions? options)
        {
            var stream = _stream!;
            if (_closeBeforeFirstRow != null)
            {
                // The header is done and <sheetData> is still an open start tag. WriteRaw of nothing closes it (and
                // is the only way to, short of writing a child), then the flush moves what the XmlWriter holds
                // into the stream, ahead of the rows. Not done earlier so that a sheet with no rows keeps the
                // self-closing <sheetData />.
                var part = _closeBeforeFirstRow;
                _closeBeforeFirstRow = null;
                part.WriteRaw(string.Empty);
                part.Flush();
            }
            _buf = stream.Buffer;
            _n = stream.Count;

            Room();
            AppendFixed(RowOpen);
            AppendDigits(_rowDigits, _rowDigitsLength);
            AppendFixed(Quote);
            if (_rowStyle != null || options != null)
            {
                // Rows with options are the exception, so these attributes are composed as strings.
                var sb = new StringBuilder();
                if (_rowStyle != null)
                    sb.Append(" s=\"").Append(_rowStyle.Id.ToString(CultureInfo.InvariantCulture)).Append("\" customFormat=\"1\"");
                if (options != null)
                {
                    if (options.Height.HasValue)
                        sb.Append(" ht=\"").Append(options.Height.Value.ToString(CultureInfo.InvariantCulture)).Append("\" customHeight=\"1\"");
                    if (options.Hidden) sb.Append(" hidden=\"1\"");
                    if (options.OutlineLevel > 0) sb.Append(" outlineLevel=\"").Append(options.OutlineLevel.ToString(CultureInfo.InvariantCulture)).Append('"');
                    if (options.Collapsed) sb.Append(" collapsed=\"1\"");
                }
                if (sb.Length > 0)
                {
                    EnsureRoom(sb.Length + 8);
                    AppendAscii(sb.ToString());
                }
            }
            AppendFixed(TagEnd);
            _rowEmpty = true;
        }

        /// <summary>The row's element written through the XmlWriter, which also writes the attributes copied from an edited row.</summary>
        private void BeginXml(RowOptions? options, IReadOnlyList<RawAttribute>? rawAttributes)
        {
            _w!.WriteStartElement("row");
            _w.WriteAttributeString("r", RowIndex.ToString(CultureInfo.InvariantCulture));
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
                    _w.WriteAttributeString("ht", options.Height.Value.ToString(CultureInfo.InvariantCulture));
                    _w.WriteAttributeString("customHeight", "1");
                }
                if (options.Hidden) _w.WriteAttributeString("hidden", "1");
                if (options.OutlineLevel > 0) _w.WriteAttributeString("outlineLevel", options.OutlineLevel.ToString(CultureInfo.InvariantCulture));
                if (options.Collapsed) _w.WriteAttributeString("collapsed", "1");
            }
            if (_buf.Length == 0) _buf = new byte[XmlTargetBufferLength];
            _n = 0;
        }

        /// <summary>Closes the <c>&lt;row&gt;</c> element if one is open.</summary>
        internal void End()
        {
            if (!IsOpen) return;
            if (_stream != null)
            {
                // After a failed flush the staged bytes are gone and the buffer holds nothing to take a byte back
                // from, so a row left open by the failure must not be closed by hand; the original error is reported.
                _stream.ThrowIfFaulted();
                if (_rowEmpty)
                {
                    // Nothing was written after the start tag's closing '>', so it is the last byte: take it back
                    // and write the empty element the way XmlWriter does, with a space before the slash.
                    _n--;
                    Room();
                    AppendFixed(EmptyEnd);
                }
                else
                {
                    Room();
                    AppendFixed(RowClose);
                }
                _stream.Count = _n;
            }
            else
            {
                // The row's cells are still in the staging buffer; they must reach the writer before its end tag.
                FlushStaged();
                _w!.WriteEndElement();
            }
            IsOpen = false;
        }

        /// <summary>
        /// Lets go of the staging memory once the sheet is sealed. Every write method checks for an open row
        /// first, so nothing can reach the empty buffer afterwards. The direct target's buffer belongs to the
        /// package's part stream and stays where it is.
        /// </summary>
        internal void ReleaseBuffer()
        {
            _buf = Array.Empty<byte>();
            _n = 0;
            _chars = null;
            _decoder = null;
            _columnStates = Array.Empty<StringColumnState>();
            _fragment = null;
            _fragmentBytes = null;
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
                BeginCellRaw(effective, null);
                // XmlWriter closes a childless element as " />" (with the space); keep the file byte for byte.
                AppendFixed(EmptyEnd);
            }
            NextColumn++;
            return this;
        }

        /// <summary>
        /// Writes a text cell. <c>null</c> leaves the column empty. Where the text is stored (in the shared
        /// string table or inside the cell) follows <see cref="PackageOptions.StringStorage"/>; by default that is
        /// decided per column from how much its values repeat. The cell reads the same either way.
        /// </summary>
        public RowWriter WriteCell(string? value, CellStyle? style = null)
        {
            if (value == null) return WriteEmpty(style);
            switch (_stringStorage)
            {
                case StringStorage.Auto: return WriteAutoString(value, style);
                case StringStorage.Shared: return WriteSharedString(value, style);
                default: return WriteInlineString(value, style);
            }
        }

        /// <summary>
        /// Writes a text cell inline in the sheet instead of through the shared string table. Use for
        /// values that are unlikely to repeat (free text, GUIDs) to keep the table small.
        /// </summary>
        public RowWriter WriteInlineString(string? value, CellStyle? style = null)
        {
            if (value == null) return WriteEmpty(style);
            EnsureOpen();
            return WriteInlineCore(value, style);
        }

        /// <summary>The inline cell itself; the caller has checked that a row is open.</summary>
        /// <param name="value">The text.</param>
        /// <param name="style">The cell style, or null.</param>
        /// <param name="shareIfChanged">
        /// Set by <see cref="StringStorage.Auto"/> for a column judged inline: text that <see cref="XmlTextUtil.Sanitize"/> has to
        /// change (control characters, a literal <c>_xHHHH_</c>) goes through the shared table instead, because some readers
        /// (openpyxl, and pandas through it) decode the <c>_xHHHH_</c> escapes in shared strings but not in inline ones.
        /// </param>
        private RowWriter WriteInlineCore(string value, CellStyle? style, bool shareIfChanged = false)
        {
            // The common cell (no style, no address, plain text) is composed in one go; anything else takes the general route.
            if (style == null && _rowStyle == null && !_alwaysWriteReference && NextColumn == _lastWrittenColumn + 1 && NextColumn <= CellReference.MaxColumns
                && TryAppendInlineText(value, dense: true))
            {
                _lastWrittenColumn = NextColumn;
                _lastStyleId = 0;
                EnsureRoom(CellReserve);
                AppendFixed(InlineClose);
                if (_cache) _sheet.RecordText(RowIndex, NextColumn);
                EndCellRaw(value);
                return this;
            }

            if (shareIfChanged && !ReferenceEquals(XmlTextUtil.Sanitize(value), value))
            {
                if (_sharedStrings!.TryGetOrAdd(XmlTextUtil.Sanitize(value), out var sharedIndex)) return WriteSharedIndex(sharedIndex, value, style);
            }

            BeginCellRaw(style, TypeInline);
            if (!TryAppendInlineText(value, dense: false))
            {
                var text = XmlTextUtil.Sanitize(value);
                AppendFixed(XmlTextUtil.NeedsSpacePreserve(text) ? InlineOpenPreserve : InlineOpen);
                AppendText(text);
            }
            EnsureRoom(CellReserve);
            AppendFixed(InlineClose);
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(value);
            return this;
        }

        /// <summary>
        /// The longest text <see cref="TryAppendInlineText"/> takes. Its worst case is six bytes per character, which has
        /// to fit in the staging buffer next to the rest of the cell without a flush in the middle: a flush could not be
        /// taken back when the text turns out to need the general path.
        /// </summary>
        private const int FastTextMaxLength = 512;

        /// <summary>
        /// Writes <c>&lt;is&gt;&lt;t&gt;</c> and the text of an inline cell in one pass over the characters, for the
        /// text that needs nothing done to it but encoding: letters, digits, punctuation, accents, any other
        /// character of the Basic Multilingual Plane, surrogate pairs, tabs and line breaks, and <c>&amp; &lt; &gt;</c>,
        /// which become entities. This is exactly what <see cref="XmlTextUtil.Sanitize"/> followed by
        /// <see cref="AppendText"/> would write for such text, without the three scans and the copy they take.
        /// Anything that needs <see cref="XmlTextUtil.Sanitize"/> to act (a control character, an unpaired surrogate,
        /// U+FFFE or U+FFFF, an underscore that opens an <c>_xHHHH_</c> sequence), a text that is too long, and a
        /// buffer too small for it leave the staging buffer as it was and return <c>false</c>, and the caller takes the
        /// general path.
        /// </summary>
        private bool TryAppendInlineText(string value, bool dense)
        {
            var length = value.Length;
            if (length > FastTextMaxLength || _buf.Length < length * 6 + CellReserve + 64) return false;
            // Room for the worst case up front: after this the loop below never flushes, so it can always start over.
            EnsureRoom(length * 6 + 64);
            var buf = _buf;
            var start = _n;
            var n = start;

            // xml:space="preserve" is decided by the first and last character, which are untouched on this path.
            var preserve = XmlTextUtil.NeedsSpacePreserve(value);
            if (dense) _rowEmpty = false; // The cell's start tag goes out below; see BeginCellRaw for why this comes before any flush.
            var open = dense ? (preserve ? DenseInlineOpenPreserve : DenseInlineOpen) : (preserve ? InlineOpenPreserve : InlineOpen);
            Buffer.BlockCopy(open, 0, buf, n, open.Length);
            n += open.Length;

            var i = 0;
#if NET8_0_OR_GREATER
            // Most text is plain ASCII: eight characters at a time, then the scalar loop takes whatever is left.
            if (length >= 8 && System.Runtime.Intrinsics.Vector128.IsHardwareAccelerated)
            {
                i = CopyPlainAscii(value, buf, n);
                n += i;
            }
#endif
            for (; i < length; i++)
            {
                var c = value[i];
                if (c < 0x80)
                {
                    if (c > '>')
                    {
                        // Letters and most punctuation, the bulk of any text. Only the underscore needs a second look.
                        if (c == '_' && XmlTextUtil.LooksLikeEscape(value, i)) return Rewind(start);
                        buf[n++] = (byte)c;
                    }
                    else if (c >= ' ')
                    {
                        if (c == '&')
                        {
                            buf[n] = (byte)'&'; buf[n + 1] = (byte)'a'; buf[n + 2] = (byte)'m'; buf[n + 3] = (byte)'p'; buf[n + 4] = (byte)';';
                            n += 5;
                        }
                        else if (c == '<' || c == '>')
                        {
                            buf[n] = (byte)'&'; buf[n + 1] = (byte)(c == '<' ? 'l' : 'g'); buf[n + 2] = (byte)'t'; buf[n + 3] = (byte)';';
                            n += 4;
                        }
                        else buf[n++] = (byte)c;
                    }
                    else if (c == '\t' || c == '\n' || c == '\r') buf[n++] = (byte)c;
                    else return Rewind(start); // Any other control character becomes an _xHHHH_ escape.
                }
                else if (c < 0x800)
                {
                    buf[n] = (byte)(0xC0 | (c >> 6));
                    buf[n + 1] = (byte)(0x80 | (c & 0x3F));
                    n += 2;
                }
                else if (c < 0xD800 || (c > 0xDFFF && c < 0xFFFE))
                {
                    buf[n] = (byte)(0xE0 | (c >> 12));
                    buf[n + 1] = (byte)(0x80 | ((c >> 6) & 0x3F));
                    buf[n + 2] = (byte)(0x80 | (c & 0x3F));
                    n += 3;
                }
                else if (c <= 0xDBFF && i + 1 < length && value[i + 1] >= 0xDC00 && value[i + 1] <= 0xDFFF)
                {
                    var scalar = 0x10000 + ((c - 0xD800) << 10) + (value[i + 1] - 0xDC00);
                    buf[n] = (byte)(0xF0 | (scalar >> 18));
                    buf[n + 1] = (byte)(0x80 | ((scalar >> 12) & 0x3F));
                    buf[n + 2] = (byte)(0x80 | ((scalar >> 6) & 0x3F));
                    buf[n + 3] = (byte)(0x80 | (scalar & 0x3F));
                    n += 4;
                    i++;
                }
                else return Rewind(start); // An unpaired surrogate, U+FFFE or U+FFFF.
            }
            _n = n;
            return true;
        }

#if NET8_0_OR_GREATER
        /// <summary>
        /// Copies the leading characters of <paramref name="value"/> that are plain ASCII (from the space to DEL, and
        /// none of <c>&amp; &lt; &gt; _</c>, which are the ones that need a look) to <paramref name="destination"/> as
        /// bytes, eight characters per step, and returns how many it copied. When the whole text is plain and at least
        /// eight characters long the last step overlaps the one before it instead of leaving a tail.
        /// </summary>
        private static int CopyPlainAscii(string value, byte[] buffer, int destination)
        {
            var length = value.Length;
            ref var source = ref System.Runtime.CompilerServices.Unsafe.As<char, ushort>(ref System.Runtime.InteropServices.MemoryMarshal.GetReference(value.AsSpan()));
            var done = 0;
            while (done + 8 <= length)
            {
                var block = System.Runtime.Intrinsics.Vector128.LoadUnsafe(ref source, (nuint)done);
                if (!IsPlainAscii(block)) return done;
                StoreNarrowed(block, buffer, destination + done);
                done += 8;
            }
            if (done < length)
            {
                var from = length - 8;
                var block = System.Runtime.Intrinsics.Vector128.LoadUnsafe(ref source, (nuint)from);
                if (IsPlainAscii(block))
                {
                    StoreNarrowed(block, buffer, destination + from);
                    return length;
                }
            }
            return done;
        }

        /// <summary>True when every one of the eight characters is from the space to DEL and none is <c>&amp;</c>, <c>&lt;</c>, <c>&gt;</c> or <c>_</c>.</summary>
        private static bool IsPlainAscii(System.Runtime.Intrinsics.Vector128<ushort> block)
        {
            // (c - 0x20) as an unsigned number is 0x60 or more exactly when c is below the space or above DEL.
            var outside = System.Runtime.Intrinsics.Vector128.GreaterThanOrEqual(block - System.Runtime.Intrinsics.Vector128.Create((ushort)0x20), System.Runtime.Intrinsics.Vector128.Create((ushort)0x60));
            var special = System.Runtime.Intrinsics.Vector128.Equals(block, System.Runtime.Intrinsics.Vector128.Create((ushort)'&'))
                | System.Runtime.Intrinsics.Vector128.Equals(block, System.Runtime.Intrinsics.Vector128.Create((ushort)'<'))
                | System.Runtime.Intrinsics.Vector128.Equals(block, System.Runtime.Intrinsics.Vector128.Create((ushort)'>'))
                | System.Runtime.Intrinsics.Vector128.Equals(block, System.Runtime.Intrinsics.Vector128.Create((ushort)'_'));
            return System.Runtime.Intrinsics.Vector128.ExtractMostSignificantBits(outside | special) == 0;
        }

        /// <summary>Writes the low byte of each of the eight characters, which are known to be ASCII, as eight bytes. The caller has reserved room.</summary>
        private static void StoreNarrowed(System.Runtime.Intrinsics.Vector128<ushort> block, byte[] buffer, int position)
        {
            var narrowed = System.Runtime.Intrinsics.Vector128.ToScalar(System.Runtime.Intrinsics.Vector128.AsUInt64(System.Runtime.Intrinsics.Vector128.Narrow(block, block)));
            System.Runtime.CompilerServices.Unsafe.WriteUnaligned(ref buffer[position], narrowed);
        }
#endif
        /// <summary>Takes back what <see cref="TryAppendInlineText"/> wrote: nothing was flushed, so the position is all there is to restore.</summary>
        private bool Rewind(int position)
        {
            _n = position;
            return false;
        }

        /// <summary>
        /// <see cref="StringStorage.Auto"/>: a column that has been judged not worth sharing pays one array read and
        /// is written inline, without hashing the text. Any other column goes through the shared table, and the
        /// outcome of that lookup (repeat or new) is what lets <see cref="StringColumnState.Observe"/> judge it from a
        /// window of cells (see <see cref="StringColumnPolicy"/>).
        /// </summary>
        private RowWriter WriteAutoString(string value, CellStyle? style)
        {
            EnsureOpen();
            var states = _columnStates;
            var column = NextColumn;
            if ((uint)column >= (uint)states.Length)
            {
                // Past Excel's last column the cell is refused by BeginCellRaw; nothing to track.
                if (column > CellReference.MaxColumns) return WriteInlineCore(value, style);
                states = GrowColumnStates(column);
            }
            ref var state = ref states[column];
            // A column judged inline, or waiting for its next look (which this call counts), pays this one test and no lookup.
            if (state.WritesInline()) return WriteInlineCore(value, style, shareIfChanged: true);
            if (!_sharedStrings!.TryGetOrAdd(XmlTextUtil.Sanitize(value), out var index, out var existed))
            {
                // The table is full (see PackageOptions.MaxSharedStringChars): a new text is written inline. It still
                // counts as a text that did not repeat, so the column is judged and can stop looking things up.
                state.Observe(hit: false, value.Length);
                return WriteInlineCore(value, style);
            }
            state.Observe(existed, value.Length);
            return WriteSharedIndex(index, value, style);
        }

        /// <summary>Makes room in <see cref="_columnStates"/> for <paramref name="column"/>, growing it to the next power of two (at most the 16,384 columns Excel has).</summary>
        private StringColumnState[] GrowColumnStates(int column)
        {
            var size = Math.Max(16, _columnStates.Length);
            while (size <= column) size *= 2;
            if (size > CellReference.MaxColumns + 1) size = CellReference.MaxColumns + 1;
            var grown = new StringColumnState[size];
            Array.Copy(_columnStates, grown, _columnStates.Length);
            return _columnStates = grown;
        }

        private RowWriter WriteSharedIndex(int index, string value, CellStyle? style)
        {
            BeginCellRaw(style, TypeSharedString);
            AppendFixed(ValueOpen);
            AppendInteger(index);
            AppendFixed(ValueClose);
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(value);
            return this;
        }

        private RowWriter WriteSharedString(string value, CellStyle? style)
        {
            EnsureOpen();
            // A new string that no longer fits the table's memory budget is written inline instead; strings
            // already in the table keep being shared.
            if (!_sharedStrings!.TryGetOrAdd(XmlTextUtil.Sanitize(value), out var index))
                return WriteInlineCore(value, style);
            return WriteSharedIndex(index, value, style);
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
            // Checked before anything is written, so a bad call leaves the row as it was.
            foreach (var run in runs)
                if (run == null) throw new ArgumentException("Runs cannot be null.", nameof(runs));
            EnsureOpen();

            // The runs are rendered first. A run can fail (a font name XML cannot carry), and by then nothing of the
            // cell may be in the row: a half-written cell would make the whole sheet unreadable.
            var rendered = RenderRuns(runs);

            // Rich text is written inline: shared strings would need the same run markup and rarely repeat.
            BeginCellRaw(style, TypeInline);
            AppendFixed(IsOpenTag);
            try
            {
                AppendBytes(rendered.GetBuffer(), (int)rendered.Length);
            }
            finally
            {
                // The rendering stream is reused by the next rich-text cell, so it is emptied even when the copy
                // fails: stale runs would otherwise be written into whatever cell comes next.
                rendered.SetLength(0);
            }
            EnsureRoom(CellReserve);
            AppendFixed(IsCloseTag);
            string? text = null;
            if (_observe)
            {
                var sb = new StringBuilder();
                foreach (var run in runs) sb.Append(run.Text);
                text = sb.ToString();
            }
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(text);
            return this;
        }

        /// <summary>
        /// Renders the runs with a real XmlWriter (they are the only markup here with enough structure to be worth
        /// it, and rich text is not a bulk case) and returns the finished UTF-8 bytes. The runs use plain element
        /// names, so the output is the same as if they had been written in place. The returned stream is reused: the
        /// caller copies the bytes out and empties it.
        /// </summary>
        private MemoryStream RenderRuns(RichText[] runs)
        {
            if (_fragment == null)
            {
                _fragmentBytes = new MemoryStream();
                _fragment = XmlWriter.Create(_fragmentBytes, new XmlWriterSettings
                {
                    Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                    CloseOutput = false,
                    ConformanceLevel = ConformanceLevel.Fragment,
                    OmitXmlDeclaration = true,
                    NewLineHandling = NewLineHandling.None,
                });
            }
            var bytes = _fragmentBytes!;
            try
            {
                foreach (var run in runs) run.WriteTo(_fragment);
                _fragment.Flush();
            }
            catch
            {
                // A writer that threw is in an unknown state; the next rich text cell starts a new one.
                _fragment = null;
                _fragmentBytes = null;
                throw;
            }
            return bytes;
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
            AppendFixed(ValueOpen);
            AppendNumber(value);
            AppendFixed(ValueClose);
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
            AppendFixed(ValueOpen);
            AppendInteger(value);
            AppendFixed(ValueClose);
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCellRaw(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell. Values beyond 2^53 lose precision in Excel, which stores all numbers as doubles.</summary>
        public RowWriter WriteCell(long value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCellRaw(style, null);
            AppendFixed(ValueOpen);
            AppendInteger(value);
            AppendFixed(ValueClose);
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, value);
            EndCellRaw(_observe ? NumberText(value) : null);
            return this;
        }

        /// <summary>Writes a numeric cell. Excel keeps 15 significant digits.</summary>
        public RowWriter WriteCell(decimal value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCellRaw(style, null);
            AppendFixed(ValueOpen);
            AppendNumber(value);
            AppendFixed(ValueClose);
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, (double)value);
            EndCellRaw(_observe ? NumberText((double)value) : null);
            return this;
        }

        /// <summary>Writes a boolean cell (shown as TRUE / FALSE).</summary>
        public RowWriter WriteCell(bool value, CellStyle? style = null)
        {
            EnsureOpen();
            BeginCellRaw(style, TypeBoolean);
            AppendFixed(value ? BooleanTrue : BooleanFalse);
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
            AppendFixed(ValueOpen);
            AppendNumber(serial);
            AppendFixed(ValueClose);
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
            AppendFixed(ValueOpen);
            AppendNumber(value.TotalDays);
            AppendFixed(ValueClose);
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
            BeginCellRaw(style, null);
            AppendFormula(text);
            AppendFixed(CellClose);
            // A later SUM over this cell cannot know its value, so it must not get one either.
            if (_cache) _sheet.RecordUnknown(RowIndex, NextColumn);
            EndCellRaw(null);
            return this;
        }

        /// <summary>
        /// Writes a formula with a precomputed numeric result that non-calculating readers will see. <c>NaN</c> and the
        /// infinities are not results Excel can store, so for them the formula is written without a result and Excel
        /// calculates it when the file is opened.
        /// </summary>
        public RowWriter WriteFormula(string formula, double cachedValue, CellStyle? style = null)
            => WriteFormulaWithNumber(PrepareFormula(formula), cachedValue, style);

        /// <summary>Shared by the explicit cached-value overload and the evaluator path; <paramref name="text"/> is already prepared.</summary>
        private RowWriter WriteFormulaWithNumber(string text, double cachedValue, CellStyle? style)
        {
            // NaN and the infinities are not results Excel can store, and the text a double formats to for them
            // differs between .NET and .NET Standard. No value is better than a wrong one: the formula stays and
            // Excel calculates it on open, the same as for a formula written without a result.
            if (double.IsNaN(cachedValue) || double.IsInfinity(cachedValue))
            {
                BeginCellRaw(style, null);
                AppendFormula(text);
                AppendFixed(CellClose);
                // A later SUM over this cell cannot know its value, so it must not get one either.
                if (_cache) _sheet.RecordUnknown(RowIndex, NextColumn);
                EndCellRaw(null);
                return this;
            }

            BeginCellRaw(style, null);
            AppendFormula(text);
            EnsureRoom(CellReserve);
            AppendFixed(ValueElementOpen);
            AppendNumber(cachedValue);
            AppendFixed(ValueElementClose);
            AppendFixed(CellClose);
            if (_cache) _sheet.RecordNumber(RowIndex, NextColumn, cachedValue);
            EndCellRaw(_observe ? NumberText(cachedValue) : null);
            return this;
        }

        /// <summary>Writes a formula with a precomputed text result.</summary>
        public RowWriter WriteFormula(string formula, string cachedValue, CellStyle? style = null)
        {
            if (cachedValue == null) throw new ArgumentNullException(nameof(cachedValue));
            var text = PrepareFormula(formula);
            BeginCellRaw(style, TypeStr);
            AppendFormula(text);
            var result = XmlTextUtil.Sanitize(cachedValue);
            EnsureRoom(CellReserve);
            // An element with no text is written as <v /> by XmlWriter.WriteElementString.
            if (result.Length == 0)
            {
                AppendFixed(ValueElementEmpty);
            }
            else
            {
                AppendFixed(ValueElementOpen);
                AppendText(result);
                EnsureRoom(CellReserve);
                AppendFixed(ValueElementClose);
            }
            AppendFixed(CellClose);
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(cachedValue);
            return this;
        }

        /// <summary>Writes a formula with a precomputed boolean result.</summary>
        public RowWriter WriteFormula(string formula, bool cachedValue, CellStyle? style = null)
        {
            var text = PrepareFormula(formula);
            BeginCellRaw(style, TypeBoolean);
            AppendFormula(text);
            EnsureRoom(CellReserve);
            AppendFixed(cachedValue ? ValueElementTrue : ValueElementFalse);
            AppendFixed(CellClose);
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(cachedValue ? "TRUE" : "FALSE");
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

        /// <summary>Appends the closing <c>&gt;</c> of the cell's start tag and its <c>&lt;f&gt;</c> element.</summary>
        private void AppendFormula(string text)
        {
            AppendFixed(FormulaOpen);
            AppendText(XmlTextUtil.Sanitize(text));
            EnsureRoom(CellReserve);
            AppendFixed(FormulaClose);
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
            // Edited workbooks are the only source of raw cells, and they write through an XmlWriter. The attributes
            // may carry namespaces whose declarations only that writer knows.
            if (_w == null) throw new InvalidOperationException("Raw cells belong to edited workbooks, whose rows are written through an XML writer.");

            if (NextColumn > CellReference.MaxColumns)
                throw new InvalidOperationException("Row already has 16384 cells, which is Excel's column limit.");
            FlushStaged();
            _w.WriteStartElement("c");
            // The rules for the address and the style are those of BeginCellRaw; keep the two in step.
            if (_alwaysWriteReference || NextColumn != _lastWrittenColumn + 1)
                _w.WriteAttributeString("r", CellReference.ColumnLetters(NextColumn) + RowIndex.ToString(CultureInfo.InvariantCulture));
            _lastWrittenColumn = NextColumn;
            var effective = style ?? _rowStyle;
            _lastStyleId = effective?.Id ?? 0;
            if (effective != null && !effective.IsDefault)
                _w.WriteAttributeString("s", effective.Id.ToString(CultureInfo.InvariantCulture));
            if (type != null) _w.WriteAttributeString("t", type);
            if (attributes != null)
                for (var i = 0; i < attributes.Count; i++) attributes[i].WriteTo(_w);
            if (innerXml.Length > 0) _w.WriteRaw(innerXml);
            _w.WriteEndElement();
            NextColumn++;
        }

        private RowWriter WriteError(string errorCode, CellStyle? style)
        {
            BeginCellRaw(style, TypeError);
            AppendFixed(ValueOpen);
            AppendText(errorCode);
            EnsureRoom(CellReserve);
            AppendFixed(ValueClose);
            if (_cache) _sheet.RecordText(RowIndex, NextColumn);
            EndCellRaw(errorCode);
            return this;
        }

        /// <summary>
        /// Approximates how a number will display under the cell's number format, for AutoFit: General text for
        /// unformatted numbers, the formatted digits plus literals (<c>$1,234.50</c>, <c>12.34%</c>) otherwise,
        /// and a value-independent stand-in for dates and times (<c>00/00/0000</c>). See <see cref="NumberDisplayFormat"/>.
        /// </summary>
        private string NumberText(double value) => _sheet.Workbook.Styles.DisplayFormat(_lastStyleId).Render(value);

        // ---- Cells composed as bytes ----------------------------------------------------------------------------------

        // The fixed pieces of a cell, as UTF-8 bytes. Their text is the text XmlWriter wrote.
        private static byte[] Bytes(string ascii)
        {
            var bytes = new byte[ascii.Length];
            for (var i = 0; i < ascii.Length; i++) bytes[i] = (byte)ascii[i];
            return bytes;
        }

        private static readonly byte[] RowOpen = Bytes("<row r=\"");
        private static readonly byte[] RowClose = Bytes("</row>");
        private static readonly byte[] Quote = Bytes("\"");
        private static readonly byte[] TagEnd = Bytes(">");
        // XmlWriter closes a childless element as " />" (with the space).
        private static readonly byte[] EmptyEnd = Bytes(" />");
        private static readonly byte[] CellOpen = Bytes("<c");
        private static readonly byte[] CellClose = Bytes("</c>");
        private static readonly byte[] AddressOpen = Bytes(" r=\"");
        private static readonly byte[] StyleOpen = Bytes(" s=\"");
        // Type attributes of the cell kinds, with their leading space.
        private static readonly byte[] TypeSharedString = Bytes(" t=\"s\"");
        private static readonly byte[] TypeInline = Bytes(" t=\"inlineStr\"");
        private static readonly byte[] TypeBoolean = Bytes(" t=\"b\"");
        private static readonly byte[] TypeStr = Bytes(" t=\"str\"");
        private static readonly byte[] TypeError = Bytes(" t=\"e\"");
        private static readonly byte[] ValueOpen = Bytes("><v>");
        private static readonly byte[] ValueClose = Bytes("</v></c>");
        private static readonly byte[] BooleanTrue = Bytes("><v>1</v></c>");
        private static readonly byte[] BooleanFalse = Bytes("><v>0</v></c>");
        private static readonly byte[] InlineOpen = Bytes("><is><t>");
        private static readonly byte[] InlineOpenPreserve = Bytes("><is><t xml:space=\"preserve\">");
        private static readonly byte[] InlineClose = Bytes("</t></is></c>");
        // The start of an unstyled inline cell that needs no address: the cell tag, its type and the text element in one piece.
        private static readonly byte[] DenseInlineOpen = Bytes("<c t=\"inlineStr\"><is><t>");
        private static readonly byte[] DenseInlineOpenPreserve = Bytes("<c t=\"inlineStr\"><is><t xml:space=\"preserve\">");
        private static readonly byte[] IsOpenTag = Bytes("><is>");
        private static readonly byte[] IsCloseTag = Bytes("</is></c>");
        private static readonly byte[] FormulaOpen = Bytes("><f>");
        private static readonly byte[] FormulaClose = Bytes("</f>");
        private static readonly byte[] ValueElementOpen = Bytes("<v>");
        private static readonly byte[] ValueElementClose = Bytes("</v>");
        private static readonly byte[] ValueElementEmpty = Bytes("<v />");
        private static readonly byte[] ValueElementTrue = Bytes("<v>1</v>");
        private static readonly byte[] ValueElementFalse = Bytes("<v>0</v>");
        private static readonly byte[] EntityAmpersand = Bytes("&amp;");
        private static readonly byte[] EntityLess = Bytes("&lt;");
        private static readonly byte[] EntityGreater = Bytes("&gt;");

        /// <summary>
        /// Appends <c>&lt;c</c>, the optional <c>r</c> and <c>s</c> attributes and <paramref name="typeAttribute"/>
        /// (already formatted, with its leading space), leaving the start tag open. The address is only written
        /// when the cell is not the next one in its row, and then as two pieces (cached column letters, then the
        /// row text computed once per row) so no per-cell string is allocated.
        /// </summary>
        private void BeginCellRaw(CellStyle? style, byte[]? typeAttribute)
        {
            if (NextColumn > CellReference.MaxColumns)
                throw new InvalidOperationException("Row already has 16384 cells, which is Excel's column limit.");

            // One check per cell guarantees room for everything fixed in it (see CellReserve).
            // The row is no longer empty from here on, also if the flush below fails: the row's '>' may already be
            // on its way out, and End must never try to take it back.
            _rowEmpty = false;
            if (_buf.Length - _n < CellReserve) FlushStaged();
            AppendFixed(CellOpen);
            // The r attribute is optional (ECMA-376 18.3.1.4): a cell without one sits in the column after the
            // previous cell of its row, or in column A when it is the first. So it is written only when this
            // cell is NOT the next one in line (after Skip, an unstyled WriteEmpty, or any other gap), which
            // keeps the position exact while the common dense row pays nothing for it, in bytes or in time.
            // PackageOptions.WriteCellReferences switches the shortcut off for readers that need every address.
            if (_alwaysWriteReference || NextColumn != _lastWrittenColumn + 1)
            {
                AppendFixed(AddressOpen);
                AppendAscii(CellReference.ColumnLetters(NextColumn));
                AppendDigits(_rowDigits, _rowDigitsLength);
                AppendFixed(Quote);
            }
            _lastWrittenColumn = NextColumn;

            var effective = style ?? _rowStyle;
            _lastStyleId = effective?.Id ?? 0;
            if (effective != null && !effective.IsDefault)
            {
                AppendFixed(StyleOpen);
                AppendInteger(effective.Id);
                AppendFixed(Quote);
            }
            if (typeAttribute != null) AppendFixed(typeAttribute);
        }

        private void EndCellRaw(string? displayText)
        {
            if (_observe && displayText != null) _sheet.Observe(RowIndex, NextColumn, displayText, _lastStyleId);
            NextColumn++;
        }

        /// <summary>Hands everything staged so far on, in order. Does nothing when the buffer is empty.</summary>
        private void FlushStaged()
        {
            var count = _n;
            if (count == 0) return;
            _n = 0;
            if (_stream != null)
            {
                _stream.Count = count;
                _stream.Drain();
                return;
            }

            // XML target: back to text for WriteRaw. An empty flush must not reach WriteRaw: it would close the open
            // start tag, and a row with no cells is written as a self-closing element.
            try
            {
                // The decoder keeps up to three bytes of an unfinished character from the previous flush, and when
                // the next flush completes it a full buffer of single-byte text decodes to one char per byte plus
                // the extra half of a surrogate pair: more chars than there are bytes in the buffer. So the char
                // buffer is sized with the decoder's own worst case for the buffer plus the carried-over bytes
                // (the same call exists on netstandard2.0), not with the byte count.
                var chars = _chars ??= new char[Utf8.GetMaxCharCount(_buf.Length + MaxCarriedOverBytes)];
                var decoder = _decoder ??= Utf8.GetDecoder();
                var length = decoder.GetChars(_buf, 0, count, chars, 0, flush: false);
                if (length > 0) _w!.WriteRaw(chars, 0, length);
            }
            catch
            {
                // The staged bytes are gone (_n was reset above) and they were part of a cell, so the sheet now has
                // a hole in it. The XML writer cannot tell, and a caller that catches the error and carries on
                // would save a malformed sheet without a word. The direct target has the same guarantee through its
                // part stream; here the package itself refuses to be saved.
                _sheet.Workbook.Package.MarkFailed();
                throw;
            }
        }

        /// <summary>Makes sure <paramref name="count"/> bytes fit, flushing the staged bytes if they do not.</summary>
        private void EnsureRoom(int count)
        {
            if (_buf.Length - _n < count) FlushStaged();
        }

        private void Room()
        {
            if (_buf.Length - _n < CellReserve) FlushStaged();
        }

        /// <summary>Appends a short constant. The caller has reserved room (see <see cref="CellReserve"/>).</summary>
        private void AppendFixed(byte[] constant)
        {
            var buf = _buf;
            var position = _n;
            for (var i = 0; i < constant.Length; i++) buf[position + i] = constant[i];
            _n = position + constant.Length;
        }

        /// <summary>Appends a short string that holds only ASCII (column letters, numbers). The caller has reserved room.</summary>
        private void AppendAscii(string text)
        {
            var buf = _buf;
            var position = _n;
            for (var i = 0; i < text.Length; i++) buf[position + i] = (byte)text[i];
            _n = position + text.Length;
        }

        private void AppendDigits(byte[] digits, int length)
        {
            var buf = _buf;
            var position = _n;
            for (var i = 0; i < length; i++) buf[position + i] = digits[i];
            _n = position + length;
        }

        /// <summary>Appends an integer in plain digits, straight into the staging buffer.</summary>
        private void AppendInteger(long value) => _n += Utf8Numbers.Format(value, _buf, _n);

        /// <summary>Appends a double the way <c>XmlWriter.WriteValue(double)</c> writes it.</summary>
        private void AppendNumber(double value) => _n += Utf8Numbers.Format(value, _buf, _n);

        /// <summary>Decimal counterpart of <see cref="AppendNumber(double)"/>.</summary>
        private void AppendNumber(decimal value) => _n += Utf8Numbers.Format(value, _buf, _n);

        /// <summary>Appends bytes that are already UTF-8, in pieces if they are more than fit.</summary>
        private void AppendBytes(byte[] source, int count)
        {
            var start = 0;
            while (count > 0)
            {
                var room = _buf.Length - _n;
                if (room == 0)
                {
                    FlushStaged();
                    room = _buf.Length;
                }
                var piece = count < room ? count : room;
                Buffer.BlockCopy(source, start, _buf, _n, piece);
                _n += piece;
                start += piece;
                count -= piece;
            }
        }

        /// <summary>
        /// Appends element text the way <c>XmlWriter.WriteString</c> writes it: <c>&amp;</c>, <c>&lt;</c> and
        /// <c>&gt;</c> become entities and everything else, including quotes, tabs and line breaks (the writer is set
        /// to leave newlines alone), is copied unchanged, as UTF-8. The text has been through
        /// <see cref="XmlTextUtil.Sanitize"/>, so it holds no character XML cannot carry. It is scanned once; runs
        /// without a special character are copied in bulk, and a string longer than the buffer goes out in pieces.
        /// </summary>
        private void AppendText(string text)
        {
#if NET8_0_OR_GREATER
            // The three characters that become entities are found with one vectorised search per piece, and the
            // text between them is transcoded in bulk, so ordinary text costs a few instructions per 16 characters.
            var rest = text.AsSpan();
            while (true)
            {
                var at = rest.IndexOfAny(EntityCharacters);
                AppendUtf8(at < 0 ? rest : rest.Slice(0, at));
                if (at < 0) return;
                var c = rest[at];
                if (_buf.Length - _n < 8) FlushStaged();
                AppendFixed(c == '&' ? EntityAmpersand : c == '<' ? EntityLess : EntityGreater);
                rest = rest.Slice(at + 1);
            }
#else
            var runStart = 0;
            for (var i = 0; i < text.Length; i++)
            {
                var c = text[i];
                // '&', '<' and '>' are all at or below '>', so almost every character is rejected by one compare.
                if (c > '>') continue;
                byte[] entity;
                if (c == '&') entity = EntityAmpersand;
                else if (c == '<') entity = EntityLess;
                else if (c == '>') entity = EntityGreater;
                else continue;

                AppendRun(text, runStart, i - runStart);
                if (_buf.Length - _n < 8) FlushStaged();
                AppendFixed(entity);
                runStart = i + 1;
            }
            AppendRun(text, runStart, text.Length - runStart);
#endif
        }

#if NET8_0_OR_GREATER
        private static readonly System.Buffers.SearchValues<char> EntityCharacters = System.Buffers.SearchValues.Create("&<>");

        /// <summary>
        /// Transcodes <paramref name="run"/> to UTF-8 straight into the staging buffer, in pieces when it does not all
        /// fit. The vectorised transcoder copies ASCII 16 or more characters at a time and never stops inside a
        /// surrogate pair, so a flush never lands inside a character. The text has been through
        /// <see cref="XmlTextUtil.Sanitize"/>, so it is valid UTF-16; should the transcoder disagree anyway, the
        /// careful loop below writes what is left (replacing what cannot be encoded), as it always did.
        /// </summary>
        private void AppendUtf8(ReadOnlySpan<char> run)
        {
            while (!run.IsEmpty)
            {
                // A character needs at most four bytes, so with 16 free bytes every call makes progress.
                if (_buf.Length - _n < 16) FlushStaged();
                var status = System.Text.Unicode.Utf8.FromUtf16(run, _buf.AsSpan(_n), out var read, out var written, replaceInvalidSequences: false, isFinalBlock: true);
                _n += written;
                run = run.Slice(read);
                if (status == System.Buffers.OperationStatus.InvalidData)
                {
                    var rest = run.ToString();
                    AppendRun(rest, 0, rest.Length);
                    return;
                }
            }
        }
#endif

        private static readonly UTF8Encoding Utf8 = new UTF8Encoding(false);

        /// <summary>
        /// Copies <paramref name="count"/> characters of <paramref name="text"/> as UTF-8, flushing whenever the buffer
        /// fills. ASCII, which is nearly all text, is a plain loop; anything else is encoded in blocks that never end
        /// inside a surrogate pair, so a flush never lands inside a character.
        /// </summary>
        private void AppendRun(string text, int start, int count)
        {
            while (count > 0)
            {
                if (_buf.Length - _n < 16) FlushStaged();
                var buf = _buf;
                var n = _n;
                var room = buf.Length - n;
                var limit = count < room ? count : room;
                var i = 0;
                for (; i < limit; i++)
                {
                    var c = text[start + i];
                    if (c >= 0x80) break;
                    buf[n + i] = (byte)c;
                }
                n += i;
                start += i;
                count -= i;
                _n = n;
                if (i == limit) continue; // Finished, or the buffer is full and the loop flushes it.

                // A character above ASCII: encode a block of what is left. A character takes at most three bytes,
                // and a surrogate pair four for two characters, so room / 3 characters always fit.
                room = buf.Length - n;
                var chars = count < room / 3 ? count : room / 3;
                if (chars > 0 && chars < count && char.IsHighSurrogate(text[start + chars - 1])) chars--;
                if (chars == 0)
                {
                    FlushStaged();
                    continue;
                }
                _n = n + Utf8.GetBytes(text, start, chars, buf, n);
                start += chars;
                count -= chars;
            }
        }

        private void EnsureOpen()
        {
            if (!IsOpen)
                throw new InvalidOperationException("No row is open. Call Worksheet.Row() before writing cells.");
        }
    }
}

