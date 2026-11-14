using System;

namespace SheetStream.Reading
{
    /// <summary>
    /// Limits that keep this library safe on files from people you do not know. An .xlsx is a zip archive, so
    /// a file of a few hundred kilobytes can unpack into gigabytes; without limits a single crafted upload
    /// could exhaust the memory of the server reading it. The defaults are far above anything Excel itself
    /// can produce and only need changing for unusual files. A file that goes over a limit is rejected with
    /// an <see cref="System.IO.InvalidDataException"/> that names the limit to raise.
    /// </summary>
    /// <remarks>
    /// <para>
    /// <see cref="ExcelReader"/> takes these directly. <see cref="ExcelPackage.OpenTemplate(string, string, PackageOptions?)"/>
    /// and <see cref="ExcelPackage.Edit(string, PackageOptions?)"/> apply the same limits to the workbook they
    /// load, through <see cref="PackageOptions.SourceLimits"/>.
    /// </para>
    /// <para>
    /// With the defaults, the most a single hostile workbook can make the library hold is about the value of
    /// <see cref="MaxLoadedPartBytes"/>, plus one row and the parser's own buffers. An endpoint that takes
    /// uploads from the public and expects ordinary data should lower both that and
    /// <see cref="MaxCellTextLength"/>; see the example.
    /// </para>
    /// </remarks>
    /// <example>
    /// <code>
    /// // An import screen that only ever expects ordinary data can be much stricter than the defaults.
    /// var options = new ReaderOptions { MaxCellTextLength = 32_767, MaxLoadedPartBytes = 32 * 1024 * 1024 };
    /// using var reader = ExcelReader.Open(upload, leaveOpen: false, options);
    /// </code>
    /// </example>
    public sealed class ReaderOptions
    {
        private int _maxCellTextLength = 1_048_576;
        private long _maxRowTextLength = 16L * 1_048_576;
        private long _maxLoadedPartBytes = 256L * 1024 * 1024;
        private long _maxStreamedPartBytes = 4L * 1024 * 1024 * 1024;

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
        /// The memory, in bytes, that one workbook may make the library keep: the shared string table, the
        /// styles, the workbook and its relationships, and any sheet loaded for editing. It is enforced twice.
        /// No such part may unpack to more than this, and what all of them occupy once parsed (a string costs
        /// two bytes a character plus a small overhead, which can be several times the size in the file) is
        /// added up and may not exceed it either. Sheets that are read or filled are streamed row by row and
        /// are not subject to this limit, so a sheet of any size still works. Default 256 MB.
        /// </summary>
        /// <remarks>
        /// The largest genuine consumer is the shared string table of a workbook with millions of distinct
        /// texts. If such a file is refused, the message names this property; raise it to suit the memory you
        /// are willing to spend on one file.
        /// </remarks>
        public long MaxLoadedPartBytes
        {
            get => _maxLoadedPartBytes;
            set => _maxLoadedPartBytes = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "The limit must be greater than zero.");
        }

        /// <summary>
        /// Largest unpacked size, in bytes, of a part that is streamed and not kept: a sheet read row by row, a
        /// template sheet being filled, a part copied from a template to the output. Such a part costs no
        /// memory however large it is, but it does cost time, and a zip archive can unpack to a thousand times
        /// its size: without a limit a 30 MB upload could keep a processor busy for minutes. Default 4 GB,
        /// more than a sheet with every one of Excel's 1,048,576 rows filled a hundred columns wide.
        /// When a template or an edit is saved, the parts that are copied unchanged may not exceed this in
        /// total either.
        /// </summary>
        /// <remarks>
        /// An endpoint that expects ordinary uploads can set this far lower, to bound the work one request can
        /// cause. Cancellation (<c>ReadRowsAsync</c> takes a token) is the other tool for that.
        /// </remarks>
        public long MaxStreamedPartBytes
        {
            get => _maxStreamedPartBytes;
            set => _maxStreamedPartBytes = value > 0 ? value : throw new ArgumentOutOfRangeException(nameof(value), value, "The limit must be greater than zero.");
        }

        /// <summary>A fresh instance with the default limits.</summary>
        internal static ReaderOptions Default => new ReaderOptions();

        /// <summary>
        /// Longest run of characters allowed without a <c>&lt;</c>, and the longest CDATA section. Neither a
        /// tag nor a text node can contain a raw <c>&lt;</c>, so this bounds the largest single token the XML
        /// parser has to hold, including an oversized attribute that the text limits cannot see. Derived from
        /// the cell limit with room for entity escapes (<c>&amp;amp;</c> is five characters for one), so text
        /// within <see cref="MaxCellTextLength"/> never trips it.
        /// </summary>
        internal long MaxTokenChars => Math.Max(64 * 1024, 12L * _maxCellTextLength);

        /// <summary>
        /// Longest single tag allowed, in characters: the element name with all of its attributes. Smaller than
        /// <see cref="MaxTokenChars"/> because a tag costs the parser far more than text of the same length (a
        /// record per attribute). The longest tags in real workbooks are range lists on conditional formats
        /// and validations, a few thousand characters; this leaves room for a value as long as a cell's text.
        /// </summary>
        internal long MaxTagChars => Math.Max(64 * 1024, _maxCellTextLength);
    }
}
