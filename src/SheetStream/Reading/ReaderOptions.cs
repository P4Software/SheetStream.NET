using System;

namespace SheetStream.Reading
{
    /// <summary>
    /// Limits that keep <see cref="ExcelReader"/> safe on files from people you do not know. An .xlsx is a zip
    /// archive, so a file of a few hundred kilobytes can unpack into gigabytes; without limits a single crafted
    /// upload could exhaust the memory of the server reading it. The defaults are far above anything Excel
    /// itself can produce and only need changing for unusual files. A file that goes over a limit is rejected
    /// with an <see cref="System.IO.InvalidDataException"/> that names the limit to raise.
    /// </summary>
    /// <example>
    /// <code>
    /// // An import screen that only ever expects short values can be stricter than the defaults.
    /// var options = new ReaderOptions { MaxCellTextLength = 32_767 };
    /// using var reader = ExcelReader.Open(upload, leaveOpen: false, options);
    /// </code>
    /// </example>
    public sealed class ReaderOptions
    {
        private int _maxCellTextLength = 1_048_576;
        private long _maxRowTextLength = 16L * 1_048_576;
        private long _maxLoadedPartBytes = 256L * 1024 * 1024;

        /// <summary>
        /// Longest text, in characters, a single cell or shared string may hold. Excel's own limit is 32,767;
        /// the default of 1,048,576 leaves room for files written by other tools.
        /// </summary>
        public int MaxCellTextLength
        {
            get => _maxCellTextLength;
            set => _maxCellTextLength = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "The limit must be greater than zero.");
        }

        /// <summary>
        /// Most text, in characters, one row may hold across all of its cells, not counting shared strings
        /// (those are stored once, however many cells use them). A row is handed to the caller as a whole, so
        /// this bounds the memory one row can take. Default 16,777,216.
        /// </summary>
        public long MaxRowTextLength
        {
            get => _maxRowTextLength;
            set => _maxRowTextLength = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "The limit must be greater than zero.");
        }

        /// <summary>
        /// Largest unpacked size, in bytes, of a part the reader keeps in memory: the shared string table, the
        /// styles, the workbook and its relationships. Sheets are streamed row by row and are not subject to
        /// this limit, so a sheet of any size still reads. Default 256 MB.
        /// </summary>
        public long MaxLoadedPartBytes
        {
            get => _maxLoadedPartBytes;
            set => _maxLoadedPartBytes = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "The limit must be greater than zero.");
        }

        /// <summary>A fresh instance with the default limits.</summary>
        internal static ReaderOptions Default => new ReaderOptions();

        /// <summary>
        /// Longest run of bytes allowed without a <c>&lt;</c>. Neither a tag nor a text node can contain a raw
        /// <c>&lt;</c>, so this bounds the largest single token the XML parser has to hold, including an
        /// oversized attribute that the text limits cannot see. Derived from the cell limit with room for
        /// multi-byte characters and entity escapes, so text within <see cref="MaxCellTextLength"/> never trips it.
        /// </summary>
        internal long MaxTokenBytes => Math.Max(64 * 1024, 12L * _maxCellTextLength);
    }
}
