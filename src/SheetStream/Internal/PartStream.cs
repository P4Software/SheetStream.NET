using System;
using System.IO;

namespace SheetStream.Internal
{
    /// <summary>
    /// The write buffer in front of a zip entry. It replaces <see cref="BufferedStream"/> for three reasons.
    /// </summary>
    /// <remarks>
    /// <list type="number">
    /// <item><description>
    /// <b>The row writer composes bytes in this buffer directly.</b> A sheet's rows are written as UTF-8 straight into
    /// <see cref="Buffer"/> (see <see cref="RowWriter"/>), so a cell is formatted once and reaches the compressor in
    /// one block, instead of being formatted as characters, encoded by the XML writer and copied again.
    /// </description></item>
    /// <item><description>
    /// <b><see cref="Flush"/> does nothing.</b> <c>XmlWriter.Flush</c> flushes the stream under it, and flushing a
    /// deflate stream forces a sync-flush block that costs compression and time. The row writer has to flush the XML
    /// writer to keep the two writers' output in order, so the flush must not travel on to the compressor.
    /// </description></item>
    /// <item><description>
    /// <b>One buffer serves every part of a package.</b> Only one part is open at a time, so the package hands the same
    /// array to each stream. A <see cref="BufferedStream"/> per part allocated 64 KB for each of the dozen small parts
    /// of a workbook.
    /// </description></item>
    /// </list>
    /// Whoever writes to <see cref="Buffer"/> directly must set <see cref="Count"/> and call <see cref="Drain"/>
    /// when the buffer is full, and must have flushed the <c>XmlWriter</c> over this stream first so that nothing it
    /// wrote earlier is still waiting in the XML writer's own buffer.
    /// <para>
    /// <b>A failure is final.</b> An <c>XmlWriter</c> that meets an I/O error is dead afterwards: every later call
    /// throws, so a caller that catches the error and carries on can never save a sheet with a hole in it. The row
    /// writer bypasses the XML writer, so this stream takes over that promise: once a write to the zip entry has
    /// failed, every later write, drain and close throws too.
    /// </para>
    /// </remarks>
    internal sealed class PartStream : Stream
    {
        private readonly Stream _inner;
        private readonly bool _leaveOpen;
        private bool _closed;
        private Exception? _fault;

        /// <summary>Wraps <paramref name="inner"/> (usually a zip entry stream) with the given buffer.</summary>
        public PartStream(Stream inner, byte[] buffer, bool leaveOpen = false)
        {
            _inner = inner;
            Buffer = buffer;
            _leaveOpen = leaveOpen;
        }

        /// <summary>The pending bytes are <c>Buffer[0..Count)</c>.</summary>
        public byte[] Buffer { get; }

        /// <summary>Number of bytes waiting in <see cref="Buffer"/>.</summary>
        public int Count { get; set; }

        /// <summary>Hands everything pending to the underlying stream and empties the buffer.</summary>
        public void Drain()
        {
            ThrowIfFaulted();
            var count = Count;
            if (count == 0) return;
            Count = 0;
            WriteInner(Buffer, 0, count);
        }

        /// <summary>Writes to the underlying stream and remembers a failure for good.</summary>
        private void WriteInner(byte[] buffer, int offset, int count)
        {
            try
            {
                _inner.Write(buffer, offset, count);
            }
            catch (Exception ex)
            {
                _fault = ex;
                throw;
            }
        }

        /// <summary>
        /// Throws an <see cref="IOException"/> that carries the original error once a write to the underlying stream has
        /// failed. The row writer calls it when a row starts or ends, so a caller that swallowed the first failure
        /// meets the same kind of error on every later row and on <c>Save</c>, whatever state the writer was left in.
        /// </summary>
        internal void ThrowIfFaulted()
        {
            if (_fault != null)
                throw new IOException("An earlier write to the workbook failed, so the part cannot be completed. " + _fault.Message, _fault);
        }

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count)
        {
            ThrowIfFaulted();
            var room = Buffer.Length - Count;
            if (count <= room)
            {
                System.Buffer.BlockCopy(buffer, offset, Buffer, Count, count);
                Count += count;
                return;
            }

            Drain();
            if (count >= Buffer.Length)
            {
                // A block as big as the buffer needs no copy.
                WriteInner(buffer, offset, count);
                return;
            }
            System.Buffer.BlockCopy(buffer, offset, Buffer, 0, count);
            Count = count;
        }

        /// <summary>
        /// Does nothing on purpose: see the remarks on the type. The bytes still reach the zip entry when the buffer
        /// fills, on <see cref="Drain"/> and when the stream is closed.
        /// </summary>
        public override void Flush() { }

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (_closed) return;
            _closed = true;
            if (disposing)
            {
                // The entry is closed whatever happened, but a failed part must not look like a finished one:
                // Drain throws when an earlier write failed, and so ends the save.
                try { Drain(); }
                finally { if (!_leaveOpen) _inner.Dispose(); }
            }
            base.Dispose(disposing);
        }

        /// <inheritdoc />
        public override bool CanRead => false;

        /// <inheritdoc />
        public override bool CanSeek => false;

        /// <inheritdoc />
        public override bool CanWrite => true;

        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
