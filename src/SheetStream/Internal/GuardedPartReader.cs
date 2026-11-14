using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace SheetStream.Internal
{
    /// <summary>
    /// Decodes the unpacked bytes of a package part into text for the XML parser and stops a crafted file from
    /// making that parser allocate without bound. It enforces four things the parser has no setting for: a cap
    /// on the total bytes of a part that is loaded into memory, a cap on the longest stretch of text without a
    /// <c>&lt;</c>, the same cap on the length of a CDATA section, and a tighter cap on the length of one tag.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The parser has to hold a tag (name plus all attributes) and a CDATA section in one piece. Well-formed XML
    /// cannot contain a raw <c>&lt;</c> inside a tag, an attribute value or a text node, so each of those sits
    /// between two <c>&lt;</c> characters and bounding the gap bounds the buffer. A CDATA section is the one
    /// place where a <c>&lt;</c> is ordinary text, so inside one the count simply keeps running until
    /// <c>]]&gt;</c>. Comments and processing instructions may also contain <c>&lt;</c>, but every reader in
    /// this library is created with both ignored, and the parser then skips them without buffering.
    /// </para>
    /// <para>
    /// A tag gets its own, smaller limit because it is far more expensive than text of the same length: the
    /// parser keeps a record for every attribute, so a tag made of a million tiny attributes costs hundreds of
    /// megabytes where a million characters of text cost two. The tag is followed from its <c>&lt;</c> to the
    /// first <c>&gt;</c> outside quotes. Only the tag still open at the end of a buffer needs following; a tag
    /// that begins and ends inside one buffer is shorter than any limit this reader accepts.
    /// </para>
    /// <para>
    /// The checks run on decoded characters, not on bytes, and that is deliberate. On bytes, a part encoded as
    /// UTF-16 could hide a long attribute behind characters such as U+3C3C, whose bytes all look like
    /// <c>&lt;</c>. Decoding here first means no encoding can disguise the text the parser will see.
    /// </para>
    /// <para>
    /// Cost: one decode (which the parser would otherwise do itself) and a few vectorised scans per buffer.
    /// </para>
    /// </remarks>
    internal sealed class GuardedPartReader : TextReader
    {
        private const int ByteBufferSize = 16 * 1024;

        // "<![CDATA[" is nine characters. Scanning stops this many characters short of the end of the buffer
        // (unless the part has ended), so a marker that starts inside the scanned region is always wholly in
        // the buffer and never has to be matched across two reads.
        private const int Holdback = 8;
        private const string CDataTail = "[CDATA[";   // What follows "<!" in a CDATA start marker.

        private readonly Stream _inner;
        private readonly string _part;
        private readonly long _maxBytes;
        private readonly bool _held;
        private readonly long _maxRun;
        private readonly long _maxTag;
        private readonly byte[] _bytes = new byte[ByteBufferSize];
        private char[] _chars = Array.Empty<char>();
        private Decoder? _decoder;

        private int _start;      // Next character to hand to the parser.
        private int _limit;      // End of the characters already checked; only those are handed out.
        private int _length;     // Characters in the buffer.
        private bool _ended;     // The inner stream is exhausted and the decoder flushed.
        private long _totalBytes;
        private long _run;       // Characters since the last markup '<' (or since the CDATA section began).
        private bool _inCData;
        private Tag _tag = Tag.Closed;   // Where the scan stands in the tag that was open when the buffer ended.
        private long _tagLength;         // Characters of that tag seen in earlier buffers.
        private InvalidDataException? _failure;   // Set by the first refusal; every later read repeats it.

        /// <summary>Position inside the tag being followed across buffers.</summary>
        private enum Tag
        {
            /// <summary>Not in a tag: the last one ended, or what follows the last <c>&lt;</c> is not a tag.</summary>
            Closed,

            /// <summary>A <c>&lt;</c> was the last character seen; the next one says what kind of markup it starts.</summary>
            Fresh,

            /// <summary>Inside a tag, outside any attribute value.</summary>
            Open,

            /// <summary>Inside an attribute value in double quotes, where <c>&gt;</c> is ordinary text.</summary>
            DoubleQuoted,

            /// <summary>Inside an attribute value in single quotes.</summary>
            SingleQuoted,
        }

        /// <param name="inner">The unpacked part; disposed with this reader.</param>
        /// <param name="part">Part name, for the error messages.</param>
        /// <param name="maxBytes">Cap on the part's total unpacked size.</param>
        /// <param name="held">
        /// True when the part is kept in memory, false when it is streamed. Only decides which limit the
        /// error message names (<c>MaxLoadedPartBytes</c> or <c>MaxStreamedPartBytes</c>).
        /// </param>
        /// <param name="maxRun">Cap on consecutive characters without a <c>&lt;</c>, and on one CDATA section.</param>
        /// <param name="maxTag">
        /// Cap on the length of one tag. Enforced to the nearest buffer, so it should be well above the buffer
        /// size of 16 KB; <see cref="Reading.ReaderOptions"/> never passes less than 64 KB.
        /// </param>
        public GuardedPartReader(Stream inner, string part, long maxBytes, bool held, long maxRun, long maxTag)
        {
            _inner = inner;
            _part = part;
            _maxBytes = maxBytes;
            _held = held;
            _maxRun = maxRun;
            _maxTag = maxTag;
        }

        public override int Read(char[] buffer, int index, int count)
        {
            if (count == 0) return 0;
            ThrowIfFailed();
            try
            {
                while (_start == _limit)
                {
                    if (_ended) return 0;
                    Compact();
                    if (_decoder == null) Begin(FillWholeBuffer());
                    else Decode(_inner.Read(_bytes, 0, _bytes.Length), 0);
                    Scan();
                }
            }
            catch (InvalidDataException ex)
            {
                _failure = ex;
                throw;
            }
            return Take(buffer, index, count);
        }

        public override async Task<int> ReadAsync(char[] buffer, int index, int count)
        {
            if (count == 0) return 0;
            ThrowIfFailed();
            try
            {
                while (_start == _limit)
                {
                    if (_ended) return 0;
                    Compact();
                    if (_decoder == null) Begin(await FillWholeBufferAsync().ConfigureAwait(false));
                    else Decode(await _inner.ReadAsync(_bytes, 0, _bytes.Length).ConfigureAwait(false), 0);
                    Scan();
                }
            }
            catch (InvalidDataException ex)
            {
                _failure = ex;
                throw;
            }
            return Take(buffer, index, count);
        }

        /// <summary>
        /// A part that was refused stays refused. A caller may read again after the first failure (a package
        /// being disposed tries to finish its sheets), and must get the same answer, not text from a buffer
        /// that was left half-checked.
        /// </summary>
        private void ThrowIfFailed()
        {
            if (_failure != null) throw new InvalidDataException(_failure.Message, _failure);
        }

        public override int Read()
        {
            var one = new char[1];
            return Read(one, 0, 1) == 0 ? -1 : one[0];
        }

        public override int Peek() => -1; // Not supported, which TextReader allows; the XML parser never peeks.

        private int Take(char[] buffer, int index, int count)
        {
            var n = Math.Min(count, _limit - _start);
            Array.Copy(_chars, _start, buffer, index, n);
            _start += n;
            return n;
        }

        /// <summary>Moves the few characters that were held back to the front of the buffer.</summary>
        private void Compact()
        {
            var keep = _length - _limit;
            if (keep > 0 && _limit > 0) Array.Copy(_chars, _limit, _chars, 0, keep);
            _length = keep;
            _start = 0;
            _limit = 0;
        }

        // The first read fills the whole byte buffer (or reaches the end of the part), so the byte order mark
        // and the XML declaration are certain to be in it when the encoding is chosen.
        private int FillWholeBuffer()
        {
            var total = 0;
            int read;
            while (total < _bytes.Length && (read = _inner.Read(_bytes, total, _bytes.Length - total)) > 0) total += read;
            return total;
        }

        private async Task<int> FillWholeBufferAsync()
        {
            var total = 0;
            int read;
            while (total < _bytes.Length && (read = await _inner.ReadAsync(_bytes, total, _bytes.Length - total).ConfigureAwait(false)) > 0) total += read;
            return total;
        }

        /// <summary>Chooses the decoder from the first bytes of the part and decodes them.</summary>
        private void Begin(int count)
        {
            var encoding = DetectEncoding(count, out var preamble);
            _decoder = encoding.GetDecoder();
            _chars = new char[encoding.GetMaxCharCount(ByteBufferSize) + Holdback + 1];
            Decode(count, preamble);
            // A part shorter than the buffer ended inside the first fill; say so now instead of on the next read.
            if (count > 0 && count < _bytes.Length) Decode(0, 0);
        }

        /// <summary>
        /// Picks the encoding the way an XML parser does: a byte order mark first, then the tell-tale zero
        /// bytes of UTF-16 without a mark, then the encoding named in the XML declaration, and UTF-8 otherwise.
        /// Every decoder throws on invalid input instead of substituting replacement characters.
        /// </summary>
        private Encoding DetectEncoding(int count, out int preamble)
        {
            var b = _bytes;
            preamble = 0;
            if (count >= 4 && b[0] == 0xFF && b[1] == 0xFE && b[2] == 0 && b[3] == 0) { preamble = 4; return new UTF32Encoding(false, false, true); }
            if (count >= 4 && b[0] == 0 && b[1] == 0 && b[2] == 0xFE && b[3] == 0xFF) { preamble = 4; return new UTF32Encoding(true, false, true); }
            if (count >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF) { preamble = 3; return new UTF8Encoding(false, true); }
            if (count >= 2 && b[0] == 0xFF && b[1] == 0xFE) { preamble = 2; return new UnicodeEncoding(false, false, true); }
            if (count >= 2 && b[0] == 0xFE && b[1] == 0xFF) { preamble = 2; return new UnicodeEncoding(true, false, true); }
            if (count >= 2 && b[0] == (byte)'<' && b[1] == 0) return new UnicodeEncoding(false, false, true);
            if (count >= 2 && b[0] == 0 && b[1] == (byte)'<') return new UnicodeEncoding(true, false, true);

            var name = DeclaredEncoding(count);
            if (name == null || name.Equals("utf-8", StringComparison.OrdinalIgnoreCase) || name.Equals("utf8", StringComparison.OrdinalIgnoreCase))
                return new UTF8Encoding(false, true);
            if (name.StartsWith("utf-16", StringComparison.OrdinalIgnoreCase) || name.StartsWith("utf-32", StringComparison.OrdinalIgnoreCase) || name.StartsWith("ucs-", StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException($"Part '{_part}' says it is encoded as {name}, but its bytes are not. The file is damaged.");
            try
            {
                return Encoding.GetEncoding(name, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException)
            {
                throw new InvalidDataException($"Part '{_part}' uses the text encoding '{name}', which is not available here. Workbook parts are normally UTF-8; open the file in Excel and save it again.", ex);
            }
        }

        /// <summary>
        /// The encoding named in the XML declaration at the start of the part, or null when there is no
        /// declaration or it names none. The declaration is plain ASCII in every encoding this can lead to.
        /// </summary>
        private string? DeclaredEncoding(int count)
        {
            var b = _bytes;
            if (count < 6 || b[0] != '<' || b[1] != '?' || b[2] != 'x' || b[3] != 'm' || b[4] != 'l') return null;
            var end = Math.Min(count, 512);
            var sb = new StringBuilder();
            for (var i = 0; i < end && b[i] != '>'; i++) sb.Append(b[i] < 0x80 ? (char)b[i] : '?');
            var text = sb.ToString();
            var at = text.IndexOf("encoding", StringComparison.Ordinal);
            if (at < 0) return null;
            at += "encoding".Length;
            while (at < text.Length && (text[at] == ' ' || text[at] == '=' || text[at] == '\t' || text[at] == '\r' || text[at] == '\n')) at++;
            if (at >= text.Length || (text[at] != '"' && text[at] != '\'')) return null;
            var close = text.IndexOf(text[at], at + 1);
            if (close < 0) return null;
            var name = text.Substring(at + 1, close - at - 1).Trim();
            // The name goes into an error message, so keep it short and free of anything but a label's characters.
            if (name.Length == 0 || name.Length > 40) return null;
            foreach (var c in name)
                if (!(char.IsLetterOrDigit(c) || c == '-' || c == '_' || c == '.' || c == ':')) return null;
            return name;
        }

        /// <summary>Counts the bytes just read against the part limit and decodes them after the held-back characters.</summary>
        private void Decode(int count, int offset)
        {
            try
            {
                if (count == 0)
                {
                    _length += _decoder!.GetChars(_bytes, 0, 0, _chars, _length, flush: true);
                    _ended = true;
                    return;
                }
                _totalBytes += count;
                if (_totalBytes > _maxBytes)
                {
                    throw new InvalidDataException(_held
                        ? $"Part '{_part}' unpacks to more than {_maxBytes:N0} bytes, which is over the limit for parts held in memory. The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxLoadedPartBytes."
                        : $"Part '{_part}' unpacks to more than {_maxBytes:N0} bytes, which is over the limit for a part that is streamed. The file is damaged or was built to waste processing time. If it is genuine, raise ReaderOptions.MaxStreamedPartBytes.");
                }
                _length += _decoder!.GetChars(_bytes, offset, count - offset, _chars, _length, flush: false);
            }
            catch (DecoderFallbackException ex)
            {
                throw new InvalidDataException($"Part '{_part}' contains bytes that are not valid text in its encoding. The file is damaged.", ex);
            }
        }

        /// <summary>
        /// Checks the characters decoded since the last call and advances <see cref="_limit"/> over them. All
        /// but the last few are checked; those wait for the next read so no marker is split.
        /// </summary>
        private void Scan()
        {
            var chars = _chars;
            var target = _ended ? _length : _length - Holdback;
            var i = _limit;
            while (i < target)
            {
                if (_inCData)
                {
                    // Look for "]]>". A section made of nothing but ']' makes this loop step one character at
                    // a time, which is still linear.
                    var close = Array.IndexOf(chars, ']', i, target - i);
                    if (close < 0)
                    {
                        Add(target - i);
                        i = target;
                        break;
                    }
                    if (close + 2 < _length && chars[close + 1] == ']' && chars[close + 2] == '>')
                    {
                        Add(close - i);
                        _inCData = false;
                        _run = 0;
                        i = close + 3;
                    }
                    else
                    {
                        Add(close + 1 - i);
                        i = close + 1;
                    }
                }
                else
                {
                    var open = FindCDataStart(chars, i, target);
                    var end = open < 0 ? target : open;
                    if (end > i) Markup(chars, i, end - i);
                    if (open < 0)
                    {
                        i = target;
                        break;
                    }
                    _run = 0;
                    _inCData = true;
                    _tag = Tag.Closed;
                    i = open + 2 + CDataTail.Length;
                }
            }
            if (i > _limit) _limit = i;
        }

        /// <summary>
        /// Index of the <c>&lt;</c> of the first <c>&lt;![CDATA[</c> that starts in [<paramref name="from"/>,
        /// <paramref name="to"/>), or -1. It searches for <c>!</c>, which is rare in a workbook, and checks
        /// around each hit, so ordinary markup costs one scan.
        /// </summary>
        private int FindCDataStart(char[] chars, int from, int to)
        {
            var searchEnd = Math.Min(to + 1, _length); // The '!' sits one past the '<' that must be before 'to'.
            var p = from + 1;
            while (p < searchEnd)
            {
                var bang = Array.IndexOf(chars, '!', p, searchEnd - p);
                if (bang < 0) return -1;
                if (chars[bang - 1] == '<' && Matches(chars, bang + 1))
                    return bang - 1;
                p = bang + 1;
            }
            return -1;
        }

        private bool Matches(char[] chars, int at)
        {
            if (at + CDataTail.Length > _length) return false;
            for (var k = 0; k < CDataTail.Length; k++)
                if (chars[at + k] != CDataTail[k]) return false;
            return true;
        }

        /// <summary>Applies the run limit to a stretch of ordinary markup, where every <c>&lt;</c> starts a tag.</summary>
        private void Markup(char[] chars, int from, int count)
        {
            var first = Array.IndexOf(chars, '<', from, count);
            if (first < 0)
            {
                Add(count);
                FollowTag(chars, from, from + count);
                return;
            }
            Add(first - from);

            var last = Array.LastIndexOf(chars, '<', from + count - 1, count);
            // A gap wholly inside this stretch can only be over the limit when the stretch itself is longer
            // than the limit, which takes a limit smaller than the buffer.
            if (count > _maxRun)
            {
                var previous = first;
                while (previous < last)
                {
                    var next = Array.IndexOf(chars, '<', previous + 1, from + count - previous - 1);
                    if (next - previous - 1 > _maxRun) ThrowRun();
                    previous = next;
                }
            }
            _run = from + count - 1 - last;
            if (_run > _maxRun) ThrowRun();

            // The last '<' starts whatever is still open when this stretch ends. A '<' cannot appear inside a
            // tag, so any tag that was being followed before it is over as far as the parser is concerned.
            _tag = Tag.Fresh;
            _tagLength = 0;
            FollowTag(chars, last + 1, from + count);
        }

        /// <summary>
        /// Continues through the tag that is open, looking for the <c>&gt;</c> that ends it, and refuses a tag
        /// that has gone on for longer than the limit. Quotes are tracked because <c>&gt;</c> is allowed inside
        /// an attribute value; without that, <c>&lt;c a="&gt;"</c> followed by a million attributes would look
        /// like a short tag and a long text.
        /// </summary>
        private void FollowTag(char[] chars, int start, int end)
        {
            if (_tag == Tag.Closed || start >= end) return;
            var i = start;
            if (_tag == Tag.Fresh)
            {
                // Comments, processing instructions and declarations are not tags: the parser skips or refuses
                // them, and their content may hold quotes that do not pair up.
                if (chars[i] == '!' || chars[i] == '?')
                {
                    _tag = Tag.Closed;
                    return;
                }
                _tag = Tag.Open;
            }
            for (; i < end; i++)
            {
                var c = chars[i];
                if (_tag == Tag.Open)
                {
                    if (c == '>')
                    {
                        _tag = Tag.Closed;
                        return;
                    }
                    if (c == '"') _tag = Tag.DoubleQuoted;
                    else if (c == '\'') _tag = Tag.SingleQuoted;
                }
                else if (_tag == Tag.DoubleQuoted)
                {
                    if (c == '"') _tag = Tag.Open;
                }
                else if (c == '\'')
                {
                    _tag = Tag.Open;
                }
            }
            _tagLength += end - start;
            if (_tagLength > _maxTag)
                throw new InvalidDataException($"Part '{_part}' contains a tag longer than {_maxTag:N0} characters. The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxCellTextLength.");
        }

        private void Add(int count)
        {
            _run += count;
            if (_run > _maxRun) ThrowRun();
        }

        private void ThrowRun()
            => throw new InvalidDataException($"Part '{_part}' contains a single value or tag longer than {_maxRun:N0} characters. The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxCellTextLength.");

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
