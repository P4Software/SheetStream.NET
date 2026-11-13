using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace SheetStream.Internal
{
    /// <summary>
    /// Read-only wrapper around the unpacked bytes of a package part that stops a crafted file from making
    /// the XML parser allocate without bound. It enforces two things the parser itself has no setting for:
    /// a cap on the total bytes of a part that is loaded into memory, and a cap on the longest run of bytes
    /// without a <c>&lt;</c>.
    /// </summary>
    /// <remarks>
    /// The second cap works because well-formed XML cannot contain a raw <c>&lt;</c> inside a tag, an
    /// attribute value or a text node, so every such token sits between two <c>&lt;</c> bytes. The parser must
    /// buffer a tag whole (name plus all attributes); bounding the gap bounds that buffer. Checking it costs
    /// two vectorised scans per buffer, which does not register next to inflating and parsing the same bytes.
    /// </remarks>
    internal sealed class GuardedPartStream : Stream
    {
        private const byte LessThan = (byte)'<';

        private readonly Stream _inner;
        private readonly string _part;
        private readonly long _maxBytes;
        private readonly long _maxRun;
        private long _total;
        private long _run;

        /// <param name="inner">The unpacked part; disposed with this stream.</param>
        /// <param name="part">Part name, for the error message.</param>
        /// <param name="maxBytes">Cap on the part's total unpacked size, or <see cref="long.MaxValue"/> for a streamed part.</param>
        /// <param name="maxRun">Cap on consecutive bytes without a <c>&lt;</c>.</param>
        public GuardedPartStream(Stream inner, string part, long maxBytes, long maxRun)
        {
            _inner = inner;
            _part = part;
            _maxBytes = maxBytes;
            _maxRun = maxRun;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => _total; set => throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            if (read > 0) Inspect(buffer, offset, read);
            return read;
        }

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            if (read > 0) Inspect(buffer, offset, read);
            return read;
        }

        /// <summary>Applies both caps to the bytes just read.</summary>
        private void Inspect(byte[] buffer, int offset, int count)
        {
            _total += count;
            if (_total > _maxBytes)
                throw new InvalidDataException($"Part '{_part}' unpacks to more than {_maxBytes:N0} bytes, which is over the limit for parts held in memory. The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxLoadedPartBytes.");

            var first = Array.IndexOf(buffer, LessThan, offset, count);
            if (first < 0)
            {
                _run += count;
                if (_run > _maxRun) ThrowRun();
                return;
            }
            if (_run + (first - offset) > _maxRun) ThrowRun();

            // A gap wholly inside this buffer can only be over the limit when the buffer itself is larger than
            // the limit, which a caller with an unusually large read buffer could arrange.
            var last = Array.LastIndexOf(buffer, LessThan, offset + count - 1, count);
            if (count > _maxRun)
            {
                var previous = first;
                while (previous < last)
                {
                    var next = Array.IndexOf(buffer, LessThan, previous + 1, offset + count - previous - 1);
                    if (next - previous - 1 > _maxRun) ThrowRun();
                    previous = next;
                }
            }
            _run = offset + count - 1 - last;
        }

        private void ThrowRun()
            => throw new InvalidDataException($"Part '{_part}' contains a single value or tag longer than {_maxRun:N0} bytes. The file is damaged or was built to exhaust memory. If it is genuine, raise ReaderOptions.MaxCellTextLength.");

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
